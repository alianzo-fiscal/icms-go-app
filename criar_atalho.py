"""
Instalador do atalho "Plataforma ICMS GO".
Usa __file__ para descobrir o diretorio do app — sem depender de argumentos
passados pelo batch (que tem problemas de encoding com acentos no caminho).
"""
import os
import shutil
import subprocess
import tempfile

# Diretório onde este script está = diretório do app
app_dir = os.path.dirname(os.path.abspath(__file__))

# Pasta local sem caracteres especiais
launcher_dir = os.path.join(os.environ['LOCALAPPDATA'], 'ICMS_GO')
os.makedirs(launcher_dir, exist_ok=True)

# 1. Copiar icone
ico_src = os.path.join(app_dir, 'ICMS360.ico')
ico_dst = os.path.join(launcher_dir, 'ICMS360.ico')
shutil.copy2(ico_src, ico_dst)
print(f'[OK] Icone copiado para: {ico_dst}')

# 2. Escrever VBS com sDir fixo apontando para o app_dir real
#    Encoding cp1252 = ANSI portugues (wscript le corretamente)
vbs_src = os.path.join(app_dir, 'Iniciar ICMS GO.vbs')
vbs_dst = os.path.join(launcher_dir, 'Iniciar ICMS GO.vbs')

with open(vbs_src, 'r', encoding='utf-8') as f:
    content = f.read()

# Substitui sDir dinamico pelo caminho absoluto do app
content = content.replace(
    'sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\\") - 1)',
    f'sDir = "{app_dir}"'
)

with open(vbs_dst, 'w', encoding='cp1252') as f:
    f.write(content)

print(f'[OK] Launcher gravado em: {vbs_dst}')

# 3. Criar atalho via PowerShell parametrizado (sem Unicode no .ps1)
ps_script = """
param($vbs, $ico)
$ws      = New-Object -ComObject WScript.