@echo off
title PCB Converter - Build Script
echo ============================================
echo  PCB Pick and Place Converter - Build
echo ============================================
echo.

REM Instalar dependencias (caso nao tenha)
echo [1/3] Verificando dependencias...
pip install openpyxl pyinstaller --quiet

echo.
echo [2/3] Gerando executavel...
pyinstaller --onefile --windowed --name "PCB Converter" --add-data "README.md;." converter_app.py

echo.
echo [3/3] Limpando arquivos temporarios...
if exist "PCB Converter.spec" del "PCB Converter.spec"
if exist "build" rmdir /s /q build

echo.
echo ============================================
echo  Executavel gerado em: dist\PCB Converter.exe
echo ============================================
pause
