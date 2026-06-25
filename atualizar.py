# coding: utf-8
"""
atualizar.py — Verifica e baixa atualizações do GitHub antes de iniciar o app.
Tenta git pull primeiro. Se git não estiver disponível, usa urllib.
"""
import sys
import os
import subprocess
import urllib.request
import json
from pathlib import Path

REPO   = "alianzo-fiscal/icms-go-app"
BRANCH = "main"
ARQUIVOS = [
    "app.py",
    "analisar_entradas.py",
    "analisar_saidas.py",
    "apuracao_3abas.py",
    "combinar_xlsx.py",
    "apuracao-icms-go/scripts/combinar_xlsx.py",
    ".streamlit/config.toml",
    "pva_monitor/pva_automacao.py",
    "pva_monitor/pva_batch.py",
    "pva_monitor/certidoes_bot.py",
    "pva_monitor/certidoes_sefaz_go.py",
    "pva_monitor/fase1_lote.py",
    "pva_monitor/pva_fase2.py",
    "pva_monitor/fase1_lote.bat",
    "pva_monitor/fase2_assinar_transmitir.bat",
    "pva_monitor/config.json",
]

BASE_DIR = Path(__file__).parent
TOKEN_FILE = BASE_DIR / "github_token.txt"


def ler_token():
    if TOKEN_FILE.exists():
        return TOKEN_FILE.read_text(encoding="utf-8").strip()
    return None


def via_git():
    """Tenta atualizar via git pull."""
    git_dir = BASE_DIR / ".git"
    if not git_dir.exists():
        return False
    try:
        result = subprocess.run(
            ["git", "pull"],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True,
            timeout=30,
        )
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def via_github_api():
    """Baixa arquivos diretamente do GitHub via API/raw URL."""
    token = ler_token()
    headers = {"User-Agent": "icms-go-app-updater"}
    if token:
        headers["Authorization"] = f"token {token}"

    # Verifica o SHA do último commit
    api_url = f"https://api.github.com/repos/{REPO}/commits/{BRANCH}"
    try:
        req = urllib.request.Request(api_url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
            sha_remoto = data["sha"][:7]
    except Exception:
        return False  # sem internet ou repo privado sem token

    # Compara com versão local
    sha_file = BASE_DIR / ".version"
    sha_local = sha_file.read_text().strip() if sha_file.exists() else ""

    if sha_local == sha_remoto:
        return True  # já está atualizado

    # Baixa cada arquivo
    raw_base = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}"
    for arq in ARQUIVOS:
        url = f"{raw_base}/{arq}"
        destino = BASE_DIR / arq
        destino.parent.mkdir(parents=True, exist_ok=True)
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as resp:
                destino.write_bytes(resp.read())
        except Exception:
            pass  # arquivo não encontrado, ignora

    sha_file.write_text(sha_remoto, encoding="utf-8")
    return True


def main():
    if via_git():
        return
    via_github_api()


if __n