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


def _fechar_popups():
    """Fecha qualquer popup Java pendente antes de prosseguir."""
    fechou = False
    def cb(hwnd, _):
        nonlocal fechou
        if not win32gui.IsWindowVisible(hwnd):
            return
        titulo = win32gui.GetWindowText(hwnd)
        for t in _TITULOS_POPUP:
            if t.lower() in titulo.lower():
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                w = right - left
                h = bottom - top
                # Verifica se e popup de erro critico ("Nao Enviar")
                textos_filhos = []
                def cb_filho(h2, _):
                    txt = win32gui.GetWindowText(h2)
                    if txt:
                        textos_filhos.append(txt.lower())
                win32gui.EnumChildWindows(hwnd, cb_filho, None)
                if any("nao enviar" in t2 or "não enviar" in t2 or "not send" in t2 for t2 in textos_filhos):
                    pyautogui.click(left + w * 0.38, top + h * 0.88)
                else:
                    pyautogui.click(left + w * 0.38, top + h * 0.94)
                logging.info(f"fechou popup: '{titulo}'")
                fechou = True
                time.sleep(0.5)
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
            return True
        logging.info(f"Abrindo PVA: {self.pva_exe}")
        subprocess.Popen([str(self.pva_exe)])
        ok = self._aguardar_pva()
        if ok:
            time.sleep(2)
            _fechar_popups()   # "Atualizar Tabelas" ou avisos iniciais
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
        """Ctrl+I -> dialogo -> cola caminho via clipboard -> Enter."""
        hwnd = self._focar_pva()
        if not hwnd:
            logging.error("PVA nao encontrado para importar arquivo")
            return False
        pyautogui.hotkey("ctrl", "i")
        time.sleep(1.5)
        # Cola o caminho no dialogo de arquivo
        pyperclip.copy(str(caminho))
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        pyautogui.press("enter")
        timeout = self.cfg.get("aguardar_importacao_segundos", 90)
        time.sleep(min(timeout, 5))
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
        pyautogui.press("enter")   # seleciona no JTable
        time.sleep(0.4)
        pyautogui.press("enter")   # confirma botao OK
        time.sleep(6)              # aguarda PVA carregar o arquivo (pode ser lento para TXTs grandes)
        return True

    def abrir_escrituracao_por_posicao(self, index: int = 0) -> bool:
        """Ctrl+A -> Home -> Down x index -> Enter x2 para abrir pela posicao na lista."""
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        pyautogui.hotkey("ctrl", "a")
        time.sleep(2)
        pyautogui.press("home")    # vai para o primeiro item da lista
        time.sleep(0.4)
        for _ in range(index):
            pyautogui.press("down")
            time.sleep(0.2)
        pyautogui.press("enter")   # seleciona no JTable
        time.sleep(0.4)
        pyautogui.press("enter")   # confirma botao OK
        time.sleep(2.5)
        return True

    def validar(self) -> bool:
        """Ctrl+V -> aguarda validacao -> fecha popup resultado."""
        hwnd = self._focar_pva()
        if not hwnd:
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

    # ── fluxos completos ─────────────────────────────────────────────────────

    def fase1_processar(self, caminho: Path) -> bool:
        """Importa + valida um arquivo TXT. Escrituracao fica aberta no PVA para Fase 2."""
        logging.info(f"[Fase1] Processando: {caminho.name}")
        self.fechar_escrituracao()
        if not self.abrir_pva():
            logging.error("Nao foi possivel abrir o PVA")
            r