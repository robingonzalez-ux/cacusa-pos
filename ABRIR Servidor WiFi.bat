@echo off
title CacusaPOS Servidor WiFi
color 0D
echo.
echo  *** CacusaPOS — Servidor WiFi + ngrok ***
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  Instalando Python...
    winget install Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
)

echo  Instalando dependencias...
pip install flask flask-cors openpyxl --quiet

echo  Iniciando servidor...
python "%~dp0CacusaPOS_Servidor.py"
pause
