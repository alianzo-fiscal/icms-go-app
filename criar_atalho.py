"""
Cria o atalho "Plataforma ICMS GO" na area de trabalho.
Chamado pelo instalar.bat com o diretorio do app como argumento.
Usa PowerShell parametrizado para evitar problemas de encoding com caminhos Unicode.
"""
import sys
import os
import subprocess
import tempfile

script_dir = sys.argv[1].rstrip('\\').rstrip('/')
vbs_path   = os.path.join(script_dir, 'Iniciar ICMS GO.vbs')
ico_path   = os.path.join(script_dir, 'ICMS360.ico')

# Script PowerShell usa parametros — nenhum caminho Unicode fica embutido no .ps1
ps_script = """
param($vbs, $ico)
$ws      = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath('Desktop')
$lnk     = Join-Path $desktop 'Plataforma ICMS GO.lnk'
$s = $ws.CreateShortcut($lnk)
$s.TargetPath    = $vbs
$s.IconLocation  = $ico
$s.Description   = 'Plataforma ICMS/GO'
$s.Save()
Write-Output "Atalho criado em: $lnk"
"""

# Escreve PS1 em UTF-8 com BOM (PowerShell le corretamente)
tmp = os.path.join(tempfile.gettempdir(), 'criar_atalho_icms.ps1')
with open(tmp, 'w', encoding='utf-8-sig') as f:
    f.write(ps_script)

result = subprocess.run(
    ['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
     '-File', tmp,
     '-vbs', vbs_path,
     '-ico', ico_path],
    capture_output=True, text=True
)

try:
    os.unlink(tmp)
except Exception:
    pass

if result.returncode == 0:
    print('[OK] ' + (result.stdout.strip() or 'Atalho criado com sucesso'))
else:
    print('[ERRO] ' + result.stderr.strip())
    sys.exit(1)
