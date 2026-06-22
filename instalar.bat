@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   Instalacao completa - Plataforma ICMS/GO
echo ============================================================
echo.

echo [1/3] Instalando dependencias Python...
pip install streamlit pandas openpyxl python-docx xlrd xlsxwriter lxml --quiet
if %errorlevel% neq 0 (
    echo ERRO: falha ao instalar dependencias. Verifique se o Python esta instalado.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

echo [2/3] Criando estrutura de pastas SPED...
call "%~dp0setup_sped.bat" /silent
echo [OK] Pastas criadas.
echo.

echo [3/3] Iniciando o aplicativo...
echo O app vai abrir no navegador em instantes...
echo Para encerrar: feche esta janela ou pressione Ctrl+C
echo.
streamlit run "%~dp0app.py"
