@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   Instalacao - Plataforma ICMS/GO
echo ============================================================
echo.

echo [1/4] Instalando dependencias Python...
pip install streamlit pandas openpyxl python-docx xlrd xlsxwriter lxml --quiet
if %errorlevel% neq 0 (
    echo ERRO: falha ao instalar dependencias.
    echo Certifique-se de que o Python esta instalado: python.org/downloads
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
set VBS_SRC=%~dp0Iniciar ICMS GO.vbs
set SHORTCUT=%USERPROFILE%\Desktop\Plataforma ICMS GO.lnk

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath = '%VBS_SRC%'; $s.IconLocation = 'shell32.dll,13'; $s.Description = 'Plataforma ICMS/GO'; $s.Save()"
echo [OK] Atalho criado na area de trabalho.
echo.

echo [4/4] Iniciando o aplicativo pela primeira vez...
echo.
echo ============================================================
echo   Instalacao concluida!
echo   Um atalho "Plataforma ICMS GO" foi criado na sua area de trabalho.
echo   Use-o para abrir o app sem precisar abrir o cmd.
echo ============================================================
echo.
wscript "%~dp0Iniciar ICMS GO.vbs"
