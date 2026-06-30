@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   Alianzo Fiscal 360 — Instalacao
echo ============================================================
echo.

:: ── 1. Python ────────────────────────────────────────────────
echo [1/5] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   ERRO: Python nao encontrado!
    echo.
    echo   1. Acesse: https://www.python.org/downloads
    echo   2. Clique em "Download Python" e instale
    echo   3. IMPORTANTE: marque "Add Python to PATH"
    echo   4. Execute este arquivo novamente
    echo.
    pause
    exit /b 1
)
python --version
echo [OK] Python encontrado.
echo.

:: ── 2. Dependencias Python ───────────────────────────────────
echo [2/5] Instalando dependencias Python...
python -m pip install --upgrade pip --quiet
python -m pip install streamlit pandas openpyxl python-docx xlrd xlsxwriter lxml playwright --quiet
if %errorlevel% neq 0 (
    echo ERRO: falha ao instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

:: ── 3. Playwright browsers ───────────────────────────────────
echo [3/5] Instalando browsers para automacao (Playwright)...
python -m playwright install chromium
echo [OK] Browser instalado.
echo.

:: ── 4. Baixar arquivos do GitHub ─────────────────────────────
echo [4/5] Baixando arquivos atualizados do GitHub...
python "%~dp0atualizar.py"
if %errorlevel% neq 0 (
    echo AVISO: nao foi possivel baixar atualizacoes. Usando arquivos locais.
)
echo [OK] Arquivos atualizados.
echo.

:: ── 5. Atalho na area de trabalho ────────────────────────────
echo [5/5] Criando atalho na area de trabalho...
python "%~dp0criar_atalho.py"
if %errorlevel% neq 0 (
    echo AVISO: nao foi possivel criar o atalho automaticamente.
    echo Crie manualmente: clique direito no desktop > Novo > Atalho
    echo Destino: %~dp0Iniciar ICMS GO.vbs
)
echo.

echo ============================================================
echo   Instalacao concluida!
echo.
echo   Atalho "Alianzo Fiscal 360" criado na area de trabalho.
echo   Clique nele para abrir o app.
echo.
echo   Link web (acesso de qualquer lugar):
echo   https://icms-go-eryemu9ww3st6mwqgktcie.streamlit.app
echo ============================================================
echo.
pause
