@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ============================================================
echo   Fase 2 - Assinar e Transmitir ao SEFAZ
echo ============================================================
echo.
python pva_fase2.py
echo.
pause
