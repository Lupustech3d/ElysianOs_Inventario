@echo off
title ElysianOS - Build
color 0B

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║     ElysianOS - Sistema de Inventário   ║
echo  ║         Build do Executavel             ║
echo  ╚══════════════════════════════════════════╝
echo.

:: Ativa o venv
call .venv\Scripts\activate

echo [1/4] Instalando dependencias...
pip install pyinstaller pillow psutil openpyxl reportlab wmi pywin32 --quiet

echo [2/4] Convertendo logo para .ico...
python -c "from PIL import Image; img = Image.open('assets/logo.png'); img.save('assets/logo.ico', format='ICO', sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])"

echo [3/4] Compilando executavel...
pyinstaller ^
    --onefile ^
    --noconsole ^
    --noupx ^
    --icon=assets/logo.ico ^
    --name=ElysianOS ^
    --add-data="assets;assets" ^
    --hidden-import=wmi ^
    --hidden-import=win32com.client ^
    --hidden-import=win32api ^
    --hidden-import=win32con ^
    --hidden-import=pywintypes ^
    --hidden-import=openpyxl ^
    --hidden-import=reportlab ^
    --hidden-import=PIL ^
    --hidden-import=psutil ^
    main.py

echo [4/4] Copiando para pasta final...
if not exist "release" mkdir release
copy dist\ElysianOS.exe release\ElysianOS.exe

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║         Build concluido!                ║
echo  ║   Arquivo: release\ElysianOS.exe        ║
echo  ╚══════════════════════════════════════════╝
echo.
pause