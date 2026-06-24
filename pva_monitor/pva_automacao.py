"""
PVA Sped Fiscal - Automacao de Validacao e Transmissao em Lote
Classe PVAAutomacao: core de automacao via pyautogui + win32gui
"""
import ctypes
import json
import logging
import subprocess
import time
from pathlib import Path

import pyautogui
import pyperclip
import win32gui
import win32con
import win32process

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# Titulos de popups Java que precisam ser fechados
_TITULOS_POPUP = [
    "Informacao", "Atencao", "Aviso", "Erro", "Sucesso",
    "Validacao", "Assinatura", "Transmissao", "Falha",
    "Information", "Warning", "Importacao", "Atualizar Tabelas",
    "Mensagem",
]

user32 = ctypes.windll.user32


def _attach_foreground(hwnd):
    """Foca janela Java contornando restricao do Windows (AttachThreadInput)."""
    fore_hwnd = win32gui.GetForegroundWindow()
    fore_tid  = user32.GetWindowThreadProcessId(fore_hwnd, None)
    our_tid   = ctypes.windll.kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(fore_tid, our_tid, True)
    user32.SetForegroundWindow(hwnd)
    user32.AttachThreadInput(fore_tid, our_tid, False)
    time.sleep(0.4)


def _encontrar_janela(titulo_parcial: str):
    """Retorna hwnd da primeira janela cujo titulo contem titulo_parcial."""
    resultado = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if titulo_parcial.lower() in t.lower():
                resultado.append(hwnd)
    win32gui.EnumWindows(cb, None)
    return resultado[0] if resultado else None


def _get_javaw_pids() -> set:
    """Retorna conjunto de PIDs de processos javaw.exe ativos (PVA)."""
    import subprocess as _sp2
    try:
        r = _sp2.run(
            ["tasklist", "/FI", "IMAGENAME eq javaw.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=5
        )
        pids = set()
        for line in r.stdout.strip().splitlines():
            parts = line.strip('"').split('","')
            if len(parts) >= 2:
                try:
                    pids.add(int(parts[1]))
                except ValueError:
                    pass
        return pids
    except Exception:
        return set()


def _fechar_popups():
    """Fecha popups Java do PVA (javaw.exe) via Escape.
    Filtra por processo para nao fechar janelas de outros aplicativos.
    """
    javaw_pids = _get_javaw_pids()
    fechou = False

    def cb(hwnd, _):
        nonlocal fechou
        if not win32gui.IsWindowVisible(hwnd):
            return
        # Verifica se a janela pertence ao processo PVA (javaw.exe)
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in javaw_pids:
                return
        except Exception:
            return
        titulo = win32gui.GetWindowText(hwnd)
        if not titulo:
            return
        for t in _TITULOS_POPUP:
            if t.lower() in titulo.lower():
                _attach_foreground(hwnd)
                time.sleep(0.4)
                if not win32gui.IsWindowVisible(hwnd):
                    fechou = True
                    break
                pyautogui.press("escape")
                logging.info(f"fechou popup PVA '{titulo}' via Escape")
                fechou = True
                time.sleep(0.6)
                break

    win32gui.EnumWindows(cb, None)
    return fechou


class PVAAutomacao:
    def __init__(self, config: dict):
        self.cfg = config
        self.pva_exe  = Path(config["pva_executavel"])
        self.titulo   = config.get("pva_titulo_janela", "EFD ICMS/IPI")

    # ── utilitarios ─────────────────────────────────────────────────────────

    def _hwnd_pva(self):
        return _encontrar_janela(self.titulo)

    def _focar_pva(self):
        hwnd = self._hwnd_pva()
        if hwnd:
            _attach_foreground(hwnd)
        return hwnd

    def _aguardar_pva(self, timeout=None):
        timeout = timeout or self.cfg.get("aguardar_pva_abrir_segundos", 40)
        inicio = time.time()
        while time.time() - inicio < timeout:
            if self._hwnd_pva():
                return True
            time.sleep(1)
        return False

    # ── acoes PVA ────────────────────────────────────────────────────────────

    def abrir_pva(self):
        if self._hwnd_pva():
            logging.info("PVA ja esta aberto")
            _fechar_popups()   # fecha popup de crash pendente antes de prosseguir
            time.sleep(0.5)
            return True
        logging.info(f"Abrindo PVA: {self.pva_exe}")
        subprocess.Popen([str(self.pva_exe)])
        ok = self._aguardar_pva()
        if ok:
            # Aguarda PVA inicializar internamente (versoesModulos, tabelas, etc.)
            # A janela aparece antes da inicializacao estar completa — precisamos esperar
            aguardar_init = self.cfg.get("aguardar_pva_inicializar_segundos", 15)
            logging.info(f"PVA aberto — aguardando inicializacao interna ({aguardar_init}s)")
            time.sleep(aguardar_init)
            # Fecha popups de startup (pode aparecer apos inicializacao)
            for _ in range(8):
                if not _fechar_popups():
                    break
                time.sleep(2.0)
            time.sleep(2)
            logging.info("PVA pronto para uso")
        return ok

    def fechar_escrituracao(self):
        """Fecha escrituracao aberta (Ctrl+F), se houver."""
        hwnd = self._focar_pva()
        if not hwnd:
            return
        pyautogui.hotkey("ctrl", "f")
        time.sleep(1.5)
        _fechar_popups()

    def importar_arquivo(self, caminho: Path) -> bool:
        """Abre menu Escrituracao Fiscal > Importar, cola caminho e confirma."""
        hwnd = self._focar_pva()
        if not hwnd:
            logging.error("PVA nao encontrado para importar arquivo")
            return False
        # Abre o menu via Alt para inicializar versoesModulos (Ctrl+I sozinho
        # dispara o binding antes da inicializacao e causa NullPointerException)
        pyautogui.press("alt")
        time.sleep(0.8)
        pyautogui.press("escape")   # fecha menu sem selecionar nada
        time.sleep(0.5)
        # Agora usa Ctrl+I — versoesModulos ja foi inicializado pelo menu
        pyautogui.hotkey("ctrl", "i")
        time.sleep(2.0)
        pyperclip.copy(str(caminho))
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        pyautogui.press("enter")
        timeout = self.cfg.get("aguardar_importacao_segundos", 90)
        time.sleep(timeout)
        _fechar_popups()
        return True

    def abrir_escrituracao_mais_recente(self) -> bool:
        """Ctrl+A -> Ctrl+End -> Enter x2 para abrir a mais recente no JTable."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "a")
        time.sleep(2)
        pyautogui.hotkey("ctrl", "end")
        time.sleep(0.6)
        pyautogui.press("enter")
        time.sleep(0.4)
        pyautogui.press("enter")
        time.sleep(6)
        return True

    def abrir_escrituracao_por_posicao(self, index: int = 0) -> bool:
        """Ctrl+A -> Home -> Down x index -> Enter x2."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "a")
        time.sleep(2)
        pyautogui.press("home")
        time.sleep(0.4)
        for _ in range(index):
            pyautogui.press("down")
            time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.4)
        pyautogui.press("enter")
        time.sleep(2.5)
        return True

    def validar(self) -> bool:
        """Ctrl+V -> aguarda validacao -> fecha popup resultado."""
        hwnd = self._focar_pva()
        if not hwnd:
            logging.error("validar: PVA nao encontrado")
            return False
        pyautogui.hotkey("ctrl", "v")
        timeout = self.cfg.get("aguardar_validacao_segundos", 20)
        time.sleep(timeout)
        _fechar_popups()
        return True

    def gerar_arquivo(self) -> bool:
        """Ctrl+G -> gera arquivo para entrega."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "g")
        time.sleep(self.cfg.get("aguardar_geracao_segundos", 30))
        _fechar_popups()
        return True

    def assinar(self) -> bool:
        """Ctrl+S -> assina com certificado digital."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "s")
        time.sleep(self.cfg.get("aguardar_assinatura_segundos", 60))
        _fechar_popups()
        return True

    def transmitir(self) -> bool:
        """Ctrl+T -> transmite ao SEFAZ."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "t")
        time.sleep(self.cfg.get("aguardar_transmissao_segundos", 120))
        _fechar_popups()
        return True

    def fechar_pva(self):
        """Fecha o PVA graciosamente (Alt+F4) para o Java salvar o banco antes de sair.
        Forceful kill (taskkill) apaga as escrituracoes importadas — evitar.
        """
        import subprocess as _sp

        # 1. Fecha popups de crash pendentes (Escape) antes de enviar Alt+F4
        for _ in range(6):
            if not _fechar_popups():
                break
            time.sleep(1.0)

        # 2. Fecha PVA graciosamente via Alt+F4 na janela principal
        hwnd = self._hwnd_pva()
        if hwnd:
            _attach_foreground(hwnd)
            time.sleep(0.4)
            pyautogui.hotkey("alt", "F4")
            time.sleep(1.5)
            # Pode aparecer dialog "Deseja sair?" — fecha com Enter
            _fechar_popups()
            pyautogui.press("enter")
            logging.info("Alt+F4 enviado ao PVA — aguardando encerramento")

        # 3. Aguarda janela desaparecer (max 20s)
        for _ in range(40):
            if not self._hwnd_pva():
                break
            time.sleep(0.5)

        # 4. Fallback: taskkill via javaw.exe (PVA roda como processo Java)
        if self._hwnd_pva():
            _sp.run(["taskkill", "/F", "/IM", "javaw.exe"], capture_output=True)
            logging.warning("PVA nao fechou graciosamente — taskkill javaw.exe usado")
            time.sleep(2)

        time.sleep(1)
        logging.info("PVA fechado com sucesso")

    # ── Fase 1: importar + validar ──────────────────────────────────────────

    def fase1_processar(self, caminho: Path) -> bool:
        """Importa e valida. PVA permanece aberto com a escrituracao ativa.
        NAO fecha a escrituracao nem o PVA — Fase 2 usa diretamente.
        """
        logging.info(f"[Fase1] {caminho.name}")
        _fechar_popups()
        if not self.abrir_pva():
            logging.error("PVA nao encontrado")
            return False
        logging.info("PVA aberto — iniciando importacao")
        if not self.importar_arquivo(caminho):
            logging.error(f"Falha ao importar: {caminho.name}")
            return False
        logging.info("Importacao concluida — abrindo escrituracao")
        if not self.abrir_escrituracao_mais_recente():
            logging.error("Falha ao abrir escrituracao")
            return False
        logging.info("Escrituracao aberta — iniciando validacao")
        ok = self.validar()
        logging.info(f"Validacao concluida: {'OK' if ok else 'ERRO'}")
        # Escrituracao permanece aberta — Fase 2 usa ela diretamente
        return ok

    # ── Fase 2: gerar + assinar + transmitir ─────────────────────────────────

    def fase2_processar(self, index: int = 0) -> bool:
        """Abre escrituracao pelo indice no PVA, gera, assina e transmite.
        Pressupoe que a escrituracao ja esta importada no banco (Fase 1).
        """
        logging.info(f"[Fase2] posicao {index}")
        self.fechar_escrituracao()
        if not self.abrir_pva():
            logging.error("PVA nao encontrado")
            return False
        if not self.abrir_escrituracao_por_posicao(index):
            logging.error(f"Falha ao abrir escrituracao posicao {index}")
            return False
        logging.info("Gerando arquivo de entrega...")
        if not self.gerar_arquivo():
            logging.error("Falha ao gerar arquivo")
            return False
        logging.info("Assinando...")
        if not self.assinar():
            logging.error("Falha na assinatura")
            return False
        logging.info("Transmitindo...")
        ok = self.transmitir()
        if not ok:
            logging.error(f"Falha na transmissao posicao {index}")
        self.fechar_escrituracao()
        return ok
