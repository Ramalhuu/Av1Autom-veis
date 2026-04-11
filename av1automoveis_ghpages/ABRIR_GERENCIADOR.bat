@echo off
title Av1Automoveis - Gerenciador de Estoque
echo.
echo  ==========================================
echo   Av1Automoveis - Gerenciador de Estoque
echo   Pirituba, Sao Paulo
echo  ==========================================
echo.
python gerenciador\gerenciador.py
if errorlevel 1 (
    echo.
    echo [ERRO] Verifique se o Python esta instalado.
    echo Execute no terminal: pip install openpyxl Pillow
    pause
)
