@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ============================================================
echo   Fase 1 - Importar e Validar em Lote no PVA
echo ============================================================
echo.
python fase1_lote.py
echo.
pause
