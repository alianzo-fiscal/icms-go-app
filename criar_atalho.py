"""
Instalador do atalho "Plataforma ICMS GO".

Estratégia:
  - Copia o ícone para %LOCALAPPDATA%\ICMS_GO\ (caminho sem caracteres especiais)
  - Grava um VBS launcher em %LOCALAPPDATA%\ICMS_GO\ que aponta diretamente para
    o diretório do app (com ã no caminho) — VBScript lida com Unicode em strings
  - Cria o atalho .lnk no Desktop apontando para o LOCALAPPDATA (sem ã)

Assim o Windows encontra o ícone e o atalho sem problemas de encoding.
"""
import sys
import os
import shutil
import subprocess
import tempfile

# Diretório do app (ex: C:\Users\...\icms-go-app)
app_dir = sys.argv[1].rstrip('\\/').rstrip('\\')

# Pasta local sem caracteres especiais
launcher_dir = os.path.join(os.environ['LOCALAPPDATA'], 'ICMS_GO')
os.makedirs(launcher_dir, exist_ok=True)

# 1. Copiar ícone
ico_src = os.path.join(app_dir, 'ICMS360.ico')
ico_dst = os.path.join(launcher_dir, 'ICMS360.ico')
shutil.copy2(ico_src, ico_dst)

# 2. Escrever VBS com sDir fixo apontando para o app_dir real
#    Encoding cp1252 = ANSI português (wscript lê corretamente)
vbs_src = os.path.join(app_dir, 'Iniciar ICMS GO.vbs')
vbs_dst = os.path.join(launcher_dir, 'Iniciar ICMS GO.vbs')

with open(vbs_src, 'r', encoding='utf-8') as f:
    content = f.read()

# Substitui o sDir dinâmico pelo caminho absoluto do app
content = content.replace(
    'sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\\") - 1)',
    f'sDir = "{app_dir}"'
)

with open(vbs_dst, 'w', encoding='cp1252') a