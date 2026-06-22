@echo off
chcp 65001 >nul

set MONITOR=%USERPROFILE%\Claude\Projects\TXT_SPED_MONITOR
set VALIDADOS=%MONITOR%\Validados
set COM_ERRO=%MONITOR%\ComErro
set PVA_DIR=%USERPROFILE%\Claude\Projects\Transformar Apuracao em Arquivo SPED ICMS\pva_monitor

if not exist "%MONITOR%"   mkdir "%MONITOR%"
if not exist "%VALIDADOS%" mkdir "%VALIDADOS%"
if not exist "%COM_ERRO%"  mkdir "%COM_ERRO%"
if not exist "%PVA_DIR%"   mkdir "%PVA_DIR%"

if "%1"=="/silent" goto :eof

echo.
echo ============================================================
echo   Pastas criadas com sucesso!
echo.
echo   %MONITOR%
echo   %MONITOR%\Validados
echo   %MONITOR%\ComErro
echo   %PVA_DIR%
echo.
echo   Proximos passos:
echo   1. Copie os scripts PVA para a pasta acima
echo      (fase1_lote.bat, pva_automacao.py, config.json etc.)
echo.
echo   2. Para instalar e abrir o app de uma vez, execute:
echo      instalar.bat
echo ============================================================
echo.
pause
