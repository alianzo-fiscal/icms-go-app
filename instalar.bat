@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   Instalacao - Plataforma ICMS/GO
echo ============================================================
echo.

echo [1/4] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ============================================================
    echo   ERRO: Python nao encontrado!
    echo.
    echo   Para instalar:
    echo   1. Acesse: https://www.python.org/downloads
    echo   2. Clique em "Download Python" e instale
    echo   3. IMPORTANTE: marque "Add Python to PATH" na instalacao
    echo   4. Depois de instalar, execute este arquivo novamente
    echo ============================================================
    echo.
    pause
    exit /b 1
)
python --version
echo [OK] Python encontrado.
echo.

echo [1/4] Instalando dependencias Python...
python -m pip install streamlit pandas openpyxl python-docx xlrd xlsxwriter lxml --quiet
if %errorlevel% neq 0 (
    echo ERRO: falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

echo [2/4] Criando pastas SPED...
call "%~dp0setup_sped.bat" /silent
echo [OK] Pastas criadas.
echo.

echo [3/4] Criando atalho na area de trabalho...
set "VBS_SRC=%~dp0Iniciar ICMS GO.vbs"
set "ICO_SRC=%~dp0ICMS360.ico"

> "%TEMP%\criar_atalho.ps1" echo $vbs = '%VBS_SRC%'
>> "%TEMP%\criar_atalho.ps1" echo $ico = '%ICO_SRC%'
>> "%TEMP%\c