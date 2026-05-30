#!/usr/bin/env python3
"""
CacusaPOS Servidor WiFi - By Taitus LLC
Corre en la PC y recibe las ventas del iPhone en tiempo real.
Escribe directamente en el Excel de OneDrive al instante.
"""

import sys, subprocess, socket, threading, json, datetime, shutil
import urllib.request, urllib.error
from pathlib import Path

# ── Auto-instalar dependencias ──
def instalar(pkg):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '--quiet'])

try:
    from flask import Flask, request, jsonify
    from flask_cors import CORS
except ImportError:
    print("Instalando Flask..."); instalar('flask'); instalar('flask-cors')
    from flask import Flask, request, jsonify
    from flask_cors import CORS

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl..."); instalar('openpyxl')
    import openpyxl

import tkinter as tk
from tkinter import scrolledtext, messagebox as tk_messagebox

# ════════════════════════════════════════
#  CONFIGURACIÓN
# ════════════════════════════════════════
BASE   = Path(r"C:\Users\titaj\OneDrive\Documents\CACUSA VENTAS")
EXCEL  = BASE / "VENTAS 4.1" / "CacusaPOS_v2.xlsm"
PORT          = 5678
NGROK_DOMAIN  = "upfront-yearbook-fascism.ngrok-free.dev"
NGROK_URL     = f"https://{NGROK_DOMAIN}"
IDS_REGISTRADOS = set()   # evita duplicados en memoria

flask_app = Flask(__name__)
CORS(flask_app)            # permite conexiones desde el iPhone

@flask_app.after_request
def _skip_ngrok_warning(response):
    response.headers['ngrok-skip-browser-warning'] = '1'
    return response

ui_log        = None       # callback para mostrar en la UI
ngrok_proc    = None       # proceso ngrok

# ════════════════════════════════════════
#  ESCRITURA EN EXCEL
# ════════════════════════════════════════
_excel_lock = threading.Lock()

def ultima_fila_con_dato(ws, col, desde):
    for r in range(ws.max_row, desde - 1, -1):
        if ws.cell(r, col).value not in (None, ''):
            return r
    return desde - 1

def escribir_venta(sale):
    with _excel_lock:
        wb = openpyxl.load_workbook(EXCEL, keep_vba=True)
        tipo = sale.get('type', 'FAC')
        tax_label = f"{sale.get('taxState','')} {sale.get('taxRate',0)}%" if sale.get('taxState') else ''

        if tipo == 'FAC' and 'VENTAS' in wb.sheetnames:
            ws  = wb['VENTAS']
            row = ultima_fila_con_dato(ws, 2, 5) + 1
            ws.cell(row,  2).value = sale.get('id')
            ws.cell(row,  3).value = 'FACTURA'
            ws.cell(row,  4).value = sale.get('client')
            ws.cell(row,  5).value = round(float(sale.get('subtotal', 0)), 2)
            ws.cell(row,  6).value = round(float(sale.get('tax', 0)), 2)
            ws.cell(row,  7).value = round(float(sale.get('total', 0)), 2)
            ws.cell(row,  8).value = sale.get('reseller', 'NO')
            ws.cell(row,  9).value = tax_label
            ws.cell(row, 10).value = sale.get('status', 'Pendiente')
            ws.cell(row, 11).value = ''
            ws.cell(row, 12).value = round(float(sale.get('usps', 0)), 2)
            ws.cell(row, 13).value = round(float(sale.get('shipping', 0)), 2)
            ws.cell(row, 14).value = sale.get('date')
            ws.cell(row, 15).value = sale.get('payMethod')

        elif tipo == 'OC' and 'ORDENES' in wb.sheetnames:
            ws   = wb['ORDENES']
            row  = ultima_fila_con_dato(ws, 2, 6) + 1
            num  = row - 5
            items_str = '|'.join([
                f"{x.get('id','')}~{x.get('name','')}~{x.get('cat','')}~"
                f"{x.get('price',0)}~{x.get('qty',1)}~0"
                for x in sale.get('items', [])
            ])
            ws.cell(row,  1).value = num
            ws.cell(row,  2).value = sale.get('id')
            ws.cell(row,  3).value = sale.get('date')
            ws.cell(row,  4).value = sale.get('client')
            ws.cell(row,  5).value = sale.get('status', 'Pendiente')
            ws.cell(row,  6).value = round(float(sale.get('subtotal', 0)), 2)
            ws.cell(row,  7).value = round(float(sale.get('tax', 0)), 2)
            ws.cell(row,  8).value = round(float(sale.get('total', 0)), 2)
            ws.cell(row,  9).value = sale.get('payMethod')
            ws.cell(row, 10).value = round(float(sale.get('shipping', 0)), 2)
            ws.cell(row, 11).value = round(float(sale.get('usps', 0)), 2)
            ws.cell(row, 12).value = tax_label
            ws.cell(row, 13).value = ''
            ws.cell(row, 14).value = items_str
            ws.cell(row, 15).value = ''

        wb.save(EXCEL)

def escribir_costo(cost):
    with _excel_lock:
        wb = openpyxl.load_workbook(EXCEL, keep_vba=True)
        if 'COSTOS' in wb.sheetnames:
            ws  = wb['COSTOS']
            row = ultima_fila_con_dato(ws, 4, 2) + 1
            ws.cell(row, 2).value = cost.get('date')
            ws.cell(row, 3).value = cost.get('type')
            ws.cell(row, 4).value = cost.get('desc')
            ws.cell(row, 5).value = round(float(cost.get('amount', 0)), 2)
            ws.cell(row, 6).value = cost.get('provider', '')
            ws.cell(row, 7).value = ''
            wb.save(EXCEL)

# ════════════════════════════════════════
#  ENDPOINTS FLASK
# ════════════════════════════════════════
@flask_app.route('/', methods=['GET'])
@flask_app.route('/app', methods=['GET'])
def serve_app():
    """Sirve la app directamente — resuelve el bloqueo HTTPS/HTTP en iPhone."""
    html_path = BASE / 'Aplicacion Movil' / 'CacusaPOS.html'
    if not html_path.exists():
        return "CacusaPOS.html no encontrado", 404
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    return content, 200, {'Content-Type': 'text/html; charset=utf-8'}

@flask_app.route('/ping', methods=['GET', 'OPTIONS'])
def ping():
    return jsonify({'ok': True, 'service': 'CacusaPOS'})

@flask_app.route('/venta', methods=['POST', 'OPTIONS'])
def recibir_venta():
    if request.method == 'OPTIONS':
        return jsonify({'ok': True})
    sale = request.get_json(force=True)
    sale_id = sale.get('id', '')

    # Evitar duplicados
    if sale_id in IDS_REGISTRADOS:
        return jsonify({'ok': True, 'msg': 'ya_registrado'})

    try:
        escribir_venta(sale)
        IDS_REGISTRADOS.add(sale_id)
        _ventas_cache['ts'] = 0   # invalidar caché para que GET /ventas refleje el cambio
        hora = datetime.datetime.now().strftime('%H:%M')
        msg  = f"✅  [{hora}]  {sale_id} · {sale.get('client','')} · ${float(sale.get('total',0)):.2f}"
        if ui_log: ui_log(msg)
        return jsonify({'ok': True})
    except PermissionError:
        err = 'El archivo Excel está abierto. Ciérralo.'
        if ui_log: ui_log(f"⚠️  {err}")
        return jsonify({'ok': False, 'error': err}), 500
    except Exception as e:
        if ui_log: ui_log(f"❌  Error: {e}")
        return jsonify({'ok': False, 'error': str(e)}), 500

@flask_app.route('/costo', methods=['POST', 'OPTIONS'])
def recibir_costo():
    if request.method == 'OPTIONS':
        return jsonify({'ok': True})
    cost = request.get_json(force=True)
    try:
        escribir_costo(cost)
        hora = datetime.datetime.now().strftime('%H:%M')
        if ui_log: ui_log(f"💰  [{hora}]  Costo: {cost.get('desc','')} · ${float(cost.get('amount',0)):.2f}")
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

def _square_api(method, path, token, body=None):
    """Helper para llamar a la API de Square desde el servidor."""
    url = f'https://connect.squareup.com{path}'
    data = json.dumps(body).encode() if body else None
    req  = urllib.request.Request(
        url, data=data,
        headers={
            'Authorization':  f'Bearer {token}',
            'Content-Type':   'application/json',
            'Square-Version': '2024-01-17',
        },
        method=method
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())

@flask_app.route('/square-link', methods=['POST', 'OPTIONS'])
def crear_link_square():
    """Genera un link de pago de Square (Payment Links API)."""
    if request.method == 'OPTIONS':
        return jsonify({'ok': True})
    data        = request.get_json(force=True)
    token       = data.get('access_token', '').strip()
    location_id = data.get('location_id', '').strip()
    amount      = data.get('amount', 0)       # en centavos
    name        = data.get('name', 'CACUSA POS')
    idem_key    = data.get('idempotency_key', str(datetime.datetime.now().timestamp()))

    if not token or not location_id:
        return jsonify({'ok': False, 'error': 'Falta access_token o location_id'}), 400

    try:
        result = _square_api('POST', '/v2/online-checkout/payment-links', token, {
            'idempotency_key': idem_key,
            'quick_pay': {
                'name':         name,
                'price_money':  {'amount': int(amount), 'currency': 'USD'},
                'location_id':  location_id
            }
        })
        url = result.get('payment_link', {}).get('url', '')
        if not url:
            return jsonify({'ok': False, 'error': 'Square no devolvió URL'}), 500
        hora = datetime.datetime.now().strftime('%H:%M')
        if ui_log: ui_log(f"🔗  [{hora}]  Link generado · {name}")
        return jsonify({'ok': True, 'url': url})
    except urllib.error.HTTPError as e:
        err = json.loads(e.read()).get('errors', [{}])
        msg = err[0].get('detail', str(e)) if err else str(e)
        return jsonify({'ok': False, 'error': msg}), 400
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

@flask_app.route('/square-locations', methods=['POST', 'OPTIONS'])
def listar_locations_square():
    """Devuelve las locations de la cuenta Square (para configuración automática)."""
    if request.method == 'OPTIONS':
        return jsonify({'ok': True})
    data  = request.get_json(force=True)
    token = data.get('access_token', '').strip()
    if not token:
        return jsonify({'ok': False, 'error': 'Falta access_token'}), 400
    try:
        result    = _square_api('GET', '/v2/locations', token)
        locations = [
            {'id': loc['id'], 'name': loc.get('name', ''), 'status': loc.get('status', '')}
            for loc in result.get('locations', [])
        ]
        return jsonify({'ok': True, 'locations': locations})
    except urllib.error.HTTPError as e:
        err = json.loads(e.read()).get('errors', [{}])
        msg = err[0].get('detail', str(e)) if err else str(e)
        return jsonify({'ok': False, 'error': msg}), 400
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '127.0.0.1'

def cargar_ids_desde_excel():
    """Pre-carga los IDs ya existentes en Excel para evitar duplicados tras reinicio."""
    try:
        if not EXCEL.exists():
            return
        wb = openpyxl.load_workbook(EXCEL, keep_vba=True, read_only=True, data_only=True)
        for sheet_name in ('VENTAS', 'ORDENES'):
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
                val = row[0]
                if val and isinstance(val, str) and ('-' in val):
                    IDS_REGISTRADOS.add(val.strip())
        wb.close()
    except Exception:
        pass   # Si el Excel no se puede leer al arrancar, no es crítico

# ════════════════════════════════════════
#  LECTURA DESDE EXCEL — sincronización bidireccional con móviles
# ════════════════════════════════════════
import time as _time
_ventas_cache = {'ts': 0, 'data': []}
_CACHE_TTL    = 30   # segundos — evita re-leer el Excel en cada request

def _parse_tax_label(lbl):
    """'FL 7%' → ('FL', 7.0)  |  '' → ('', 0.0)"""
    lbl = str(lbl or '').strip()
    if not lbl:
        return '', 0.0
    parts = lbl.split()
    if len(parts) >= 2:
        try:
            return parts[0], float(parts[1].replace('%', ''))
        except Exception:
            pass
    return lbl, 0.0

def _parse_date(val):
    """Convierte datetime o string a 'DD/MM/YYYY'."""
    if hasattr(val, 'strftime'):
        return val.strftime('%d/%m/%Y')
    return str(val or '').strip()

def leer_ventas_de_excel():
    """
    Lee todas las ventas (FAC) y órdenes (OC) del Excel y las devuelve como
    lista de dicts compatibles con el formato localStorage de la app.
    Cachea 30 s para no bloquear el servidor en cada request.
    """
    now = _time.time()
    if now - _ventas_cache['ts'] < _CACHE_TTL:
        return _ventas_cache['data']

    ventas = []
    if not EXCEL.exists():
        _ventas_cache.update({'ts': now, 'data': ventas})
        return ventas

    try:
        wb = openpyxl.load_workbook(EXCEL, keep_vba=True, read_only=True, data_only=True)

        # ── Hoja VENTAS (FAC) ──
        # Col: B=id, C=tipo, D=cliente, E=subtotal, F=tax, G=total,
        #      H=reseller, I=tax_lbl, J=status, K=vacío,
        #      L=usps, M=shipping, N=fecha, O=payMethod
        if 'VENTAS' in wb.sheetnames:
            for row in wb['VENTAS'].iter_rows(min_row=5, values_only=True):
                sid = row[1]
                if not sid or '-' not in str(sid):
                    continue
                ts, tr = _parse_tax_label(row[8])
                ventas.append({
                    'id':        str(sid).strip(),
                    'type':      'FAC',
                    'client':    str(row[3]  or ''),
                    'date':      _parse_date(row[13]),
                    'subtotal':  round(float(row[4]  or 0), 2),
                    'tax':       round(float(row[5]  or 0), 2),
                    'total':     round(float(row[6]  or 0), 2),
                    'reseller':  str(row[7]  or 'NO'),
                    'taxState':  ts,  'taxRate': tr,
                    'status':    str(row[9]  or 'Pagado'),
                    'usps':      round(float(row[11] or 0), 2),
                    'shipping':  round(float(row[12] or 0), 2),
                    'payMethod': str(row[14] or ''),
                    'items': [], 'dije': 0, 'dijeQty': 0, 'ccFee': 0,
                })

        # ── Hoja ORDENES (OC) ──
        # Col: A=num, B=id, C=fecha, D=cliente, E=status,
        #      F=subtotal, G=tax, H=total, I=payMethod,
        #      J=shipping, K=usps, L=tax_lbl, M=vacío,
        #      N=items_str, O=vacío
        if 'ORDENES' in wb.sheetnames:
            for row in wb['ORDENES'].iter_rows(min_row=6, values_only=True):
                sid = row[1]
                if not sid or '-' not in str(sid):
                    continue
                ts, tr = _parse_tax_label(row[11])
                # Reconstruir items desde "id~name~cat~price~qty~0|..."
                items = []
                for seg in str(row[13] or '').split('|'):
                    p = seg.split('~')
                    if len(p) >= 5:
                        try:
                            items.append({
                                'id':    p[0], 'name': p[1], 'cat': p[2],
                                'price': float(p[3]), 'qty': int(float(p[4])),
                            })
                        except Exception:
                            pass
                ventas.append({
                    'id':        str(sid).strip(),
                    'type':      'OC',
                    'client':    str(row[3]  or ''),
                    'date':      _parse_date(row[2]),
                    'status':    str(row[4]  or 'Pendiente'),
                    'subtotal':  round(float(row[5]  or 0), 2),
                    'tax':       round(float(row[6]  or 0), 2),
                    'total':     round(float(row[7]  or 0), 2),
                    'payMethod': str(row[8]  or ''),
                    'shipping':  round(float(row[9]  or 0), 2),
                    'usps':      round(float(row[10] or 0), 2),
                    'taxState':  ts,  'taxRate': tr,
                    'items':     items,
                    'dije': 0, 'dijeQty': 0, 'ccFee': 0, 'reseller': 'NO',
                })

        wb.close()
    except Exception:
        pass   # Si el Excel está bloqueado, devuelve caché anterior

    _ventas_cache.update({'ts': _time.time(), 'data': ventas})
    return ventas


@flask_app.route('/ventas', methods=['GET', 'OPTIONS'])
def obtener_ventas():
    """Devuelve todas las ventas/órdenes del Excel para sincronizar los móviles."""
    if request.method == 'OPTIONS':
        return jsonify({'ok': True})
    try:
        ventas = leer_ventas_de_excel()
        return jsonify({'ok': True, 'ventas': ventas, 'total': len(ventas)})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


def run_flask():
    cargar_ids_desde_excel()
    flask_app.run(host='0.0.0.0', port=PORT, debug=False, use_reloader=False, threaded=True)

def run_ngrok():
    """Lanza ngrok con dominio estático en segundo plano."""
    global ngrok_proc
    ngrok_exe = shutil.which('ngrok')
    if not ngrok_exe:
        if ui_log: ui_log("⚠️  ngrok no encontrado — solo modo WiFi local")
        return
    try:
        ngrok_proc = subprocess.Popen(
            [ngrok_exe, 'http', f'--domain={NGROK_DOMAIN}', str(PORT)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        if ui_log: ui_log(f"🌐  ngrok activo → {NGROK_URL}")
    except Exception as e:
        if ui_log: ui_log(f"⚠️  ngrok error: {e}")

# ════════════════════════════════════════
#  INTERFAZ GRÁFICA
# ════════════════════════════════════════
class ServerUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CacusaPOS · Servidor WiFi")
        self.geometry("540x560")
        self.resizable(False, False)
        self.configure(bg='#FDE8F0')
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.local_ip = get_local_ip()
        self._build()
        self._start_server()

    def _build(self):
        # ── Header ──
        hdr = tk.Frame(self, bg='#C9637A', pady=16)
        hdr.pack(fill='x')
        tk.Label(hdr, text="💎  CACUSA POS",
                 font=('Helvetica', 19, 'bold'), bg='#C9637A', fg='white').pack()
        tk.Label(hdr, text="Servidor WiFi · Recibe ventas del iPhone en tiempo real",
                 font=('Helvetica', 10), bg='#C9637A', fg='#FFD6E0').pack(pady=(2, 0))

        body = tk.Frame(self, bg='#FDE8F0', padx=20, pady=14)
        body.pack(fill='both', expand=True)

        # ── Tarjeta HTTPS (ngrok — URL principal) ──
        https_card = tk.Frame(body, bg='#E8F5E9', relief='solid', bd=1)
        https_card.pack(fill='x', pady=(0, 8))

        tk.Label(https_card,
                 text="🔒  URL HTTPS — úsala en el iPhone (funciona en cualquier lugar):",
                 font=('Helvetica', 10, 'bold'), bg='#E8F5E9', fg='#1B5E20',
                 justify='left').pack(padx=14, pady=(10, 2), anchor='w')

        https_row = tk.Frame(https_card, bg='#E8F5E9')
        https_row.pack(fill='x', padx=14, pady=(0, 4))

        self.ngrok_var = tk.StringVar(value=NGROK_URL)
        tk.Label(https_row, textvariable=self.ngrok_var,
                 font=('Courier', 13, 'bold'), bg='#E8F5E9', fg='#2E7D32').pack(side='left')

        tk.Button(https_row, text="  Copiar  ",
                  font=('Helvetica', 10, 'bold'), bg='#2E7D32', fg='white',
                  activebackground='#1B5E20', relief='flat', padx=10, pady=4,
                  cursor='hand2', command=self._copiar_ngrok).pack(side='right')

        self.lbl_ngrok = tk.Label(https_card,
                 text="⏳  Iniciando ngrok...",
                 font=('Helvetica', 9), bg='#E8F5E9', fg='#558B2F')
        self.lbl_ngrok.pack(padx=14, pady=(0, 8), anchor='w')

        # ── Tarjeta WiFi local (respaldo) ──
        ip_card = tk.Frame(body, bg='white', relief='solid', bd=1)
        ip_card.pack(fill='x', pady=(0, 10))

        tk.Label(ip_card,
                 text="📶  WiFi local (solo cuando estés en casa con el servidor abierto):",
                 font=('Helvetica', 9), bg='white', fg='#A0607A',
                 justify='left').pack(padx=14, pady=(8, 2), anchor='w')

        ip_row = tk.Frame(ip_card, bg='white')
        ip_row.pack(fill='x', padx=14, pady=(0, 10))

        self.ip_var = tk.StringVar(value=f"http://{self.local_ip}:{PORT}")
        tk.Label(ip_row, textvariable=self.ip_var,
                 font=('Courier', 13, 'bold'), bg='white', fg='#C9637A').pack(side='left')

        tk.Button(ip_row, text="  Copiar  ",
                  font=('Helvetica', 10, 'bold'), bg='#C9637A', fg='white',
                  activebackground='#A04060', relief='flat', padx=10, pady=4,
                  cursor='hand2', command=self._copiar_ip).pack(side='right')

        # ── Estado ──
        self.lbl_estado = tk.Label(body,
                 text="⏳  Iniciando servidor...",
                 font=('Helvetica', 11), bg='#FDE8F0', fg='#A0607A')
        self.lbl_estado.pack(anchor='w', pady=(0, 8))

        # ── Archivo Excel ──
        excel_ok = EXCEL.exists()
        excel_color = '#4CAF7D' if excel_ok else '#E05555'
        excel_txt   = f"✅  Excel encontrado: {EXCEL.name}" if excel_ok else f"⚠️  No se encontró: {EXCEL}"
        tk.Label(body, text=excel_txt,
                 font=('Helvetica', 10), bg='#FDE8F0', fg=excel_color,
                 wraplength=480, justify='left').pack(anchor='w', pady=(0, 10))

        # ── Log ──
        tk.Label(body, text="Registro de ventas recibidas:",
                 font=('Helvetica', 10, 'bold'), bg='#FDE8F0', fg='#3A1020').pack(anchor='w', pady=(0, 4))

        log_frame = tk.Frame(body, bg='white', relief='solid', bd=1)
        log_frame.pack(fill='both', expand=True)

        self.log_txt = scrolledtext.ScrolledText(log_frame,
                 font=('Helvetica', 11), bg='white', fg='#3A1020',
                 relief='flat', bd=0, state='disabled', wrap='word')
        self.log_txt.pack(fill='both', expand=True, padx=8, pady=8)

        # ── Pie ──
        tk.Label(self,
                 text="Mantén esta ventana abierta mientras usas la app en el iPhone  ·  Las ventas se guardan al instante en tu Excel",
                 font=('Helvetica', 8), bg='#FDE8F0', fg='#C898AA',
                 wraplength=520).pack(pady=(0, 8))

    def _agregar_log(self, msg):
        self.log_txt.config(state='normal')
        self.log_txt.insert('end', msg + '\n')
        self.log_txt.see('end')
        self.log_txt.config(state='disabled')

    def _copiar_ngrok(self):
        self.clipboard_clear()
        self.clipboard_append(self.ngrok_var.get())
        self.lbl_estado.config(text="✅  URL HTTPS copiada. Ábrela en Safari del iPhone", fg='#2E7D32')
        self.after(3000, lambda: self.lbl_estado.config(
            text="🟢  Servidor activo · ngrok HTTPS listo", fg='#4CAF7D'))

    def _copiar_ip(self):
        self.clipboard_clear()
        self.clipboard_append(self.ip_var.get())
        self.lbl_estado.config(text="✅  Dirección copiada. Pégala en Ajustes → Dirección del PC", fg='#4CAF7D')
        self.after(3000, lambda: self.lbl_estado.config(
            text="🟢  Servidor activo · Esperando ventas del iPhone", fg='#4CAF7D'))

    def _start_server(self):
        global ui_log
        ui_log = lambda msg: self.after(0, self._agregar_log, msg)
        # Arrancar Flask
        t = threading.Thread(target=run_flask, daemon=True)
        t.start()
        # Arrancar ngrok después de 2s (Flask necesita estar listo)
        self.after(2000, self._start_ngrok)
        self.after(1500, lambda: self.lbl_estado.config(
            text="🟢  Servidor activo · Iniciando ngrok...", fg='#4CAF7D'))

    def _start_ngrok(self):
        threading.Thread(target=self._ngrok_worker, daemon=True).start()

    def _ngrok_worker(self):
        run_ngrok()
        # Verificar que ngrok esté respondiendo (esperar hasta 10s)
        import urllib.request, time
        for _ in range(10):
            time.sleep(1)
            try:
                with urllib.request.urlopen('http://localhost:4040/api/tunnels', timeout=1) as r:
                    data = json.loads(r.read())
                    tunnels = data.get('tunnels', [])
                    if tunnels:
                        url = tunnels[0].get('public_url', NGROK_URL)
                        self.after(0, lambda u=url: self.lbl_ngrok.config(
                            text=f"✅  ngrok activo · HTTPS listo", fg='#2E7D32'))
                        self.after(0, lambda: self.lbl_estado.config(
                            text="🟢  Servidor activo · ngrok HTTPS listo", fg='#4CAF7D'))
                        return
            except:
                pass
        # Si no se pudo verificar, igual mostrar como activo
        self.after(0, lambda: self.lbl_ngrok.config(
            text="✅  ngrok iniciado", fg='#2E7D32'))

    def on_close(self):
        if tk_messagebox.askokcancel("Cerrar", "¿Cerrar el servidor?\nLas ventas del iPhone ya no se guardarán en Excel hasta que lo abras de nuevo."):
            if ngrok_proc:
                try: ngrok_proc.terminate()
                except: pass
            self.destroy()

if __name__ == '__main__':
    ui = ServerUI()
    ui.mainloop()
