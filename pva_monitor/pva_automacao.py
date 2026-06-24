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

    # ── Utilitarios para operacoes em lote ──────────────────────────────────

    def _aguardar_dialogo_pva(self, titulo_parcial: str, timeout: int = 20) -> int:
        """Aguarda dialogo do PVA cujo titulo contem titulo_parcial.
        Filtra por processo javaw.exe para evitar janelas de outros apps.
        Retorna hwnd do dialogo ou 0 se timeout.
        """
        javaw_pids = _get_javaw_pids()
        inicio = time.time()
        while time.time() - inicio < timeout:
            encontrado = [0]

            def cb(hwnd, _):
                if encontrado[0]:
                    return
                if not win32gui.IsWindowVisible(hwnd):
                    return
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid not in javaw_pids:
                        return
                except Exception:
                    return
                t = win32gui.GetWindowText(hwnd)
                if titulo_parcial.lower() in t.lower():
                    encontrado[0] = hwnd

            win32gui.EnumWindows(cb, None)
            if encontrado[0]:
                return encontrado[0]
            time.sleep(0.5)
        return 0

    def _click_toolbar_btn(self, btn_index: int) -> bool:
        """Clica no botao da toolbar principal do PVA pelo indice (0-based).
        Posicoes estimadas: toolbar y≈55px do topo, botoes a partir de x≈14, espacamento≈27px.
        """
        hwnd = self._focar_pva()
        if not hwnd:
            return False
        rect = win32gui.GetWindowRect(hwnd)
        btn_x = rect[0] + 14 + btn_index * 27
        btn_y = rect[1] + 55
        pyautogui.click(btn_x, btn_y)
        time.sleep(0.6)
        return True

    def _confirmar_dialogo_batch(self, dlg_hwnd: int) -> bool:
        """Em dialogo de lista PVA (Verificar/Gerar/Assinar/Transmitir):
        foca o dialogo, clica na tabela, seleciona tudo (Ctrl+A) e clica OK.
        O botao OK fica em ~45% da largura e ~90% da altura do dialogo.
        """
        if not dlg_hwnd:
            return False
        _attach_foreground(dlg_hwnd)
        time.sleep(0.6)

        rect = win32gui.GetWindowRect(dlg_hwnd)
        w = rect[2] - rect[0]
        h = rect[3] - rect[1]

        # Clica na area da tabela para dar foco a ela
        pyautogui.click(rect[0] + w // 2, rect[1] + h // 2)
        time.sleep(0.4)

        # Seleciona todos os registros da tabela
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)

        # Clica no botao OK (estimado: ~45% da largura, ~90% da altura)
        ok_x = rect[0] + int(w * 0.45)
        ok_y = rect[1] + int(h * 0.90)
        logging.info(
            f"Clicando OK em ({ok_x}, {ok_y}) rect={rect} "
            f"dialogo='{win32gui.GetWindowText(dlg_hwnd)}'"
        )
        pyautogui.click(ok_x, ok_y)
        time.sleep(0.4)
        # Fallback: Enter aciona o botao default do dialogo caso o click tenha errado
        pyautogui.press("enter")
        time.sleep(0.5)
        return True

    def _batch_operacao(
        self,
        shortcut_keys: list,
        titulo_dialogo: str,
        timeout_chave: str,
        timeout_padrao: int,
        toolbar_btn_index: int = None,
    ) -> bool:
        """Executa uma operacao em lote no PVA:
        1. Envia atalho de teclado para abrir dialogo de selecao
        2. Se dialogo nao aparecer, tenta clicar no botao da toolbar (fallback)
        3. Seleciona todos os registros e confirma
        4. Aguarda conclusao (timeout configuravel)
        5. Fecha popups de resultado
        """
        hwnd = self._focar_pva()
        if not hwnd:
            logging.error("_batch_operacao: PVA nao encontrado")
            return False

        # 1. Tenta via atalho de teclado
        logging.info(f"Enviando atalho {'+'.join(shortcut_keys)}")
        pyautogui.hotkey(*shortcut_keys)
        time.sleep(1.5)

        # 2. Aguarda dialogo de selecao aparecer
        dlg = self._aguardar_dialogo_pva(titulo_dialogo, timeout=20)

        # 3. Fallback: clica no botao da toolbar
        if not dlg and toolbar_btn_index is not None:
            logging.info(
                f"Dialogo '{titulo_dialogo}' nao detectado via atalho — "
                f"tentando toolbar btn {toolbar_btn_index}"
            )
            self._click_toolbar_btn(toolbar_btn_index)
            dlg = self._aguardar_dialogo_pva(titulo_dialogo, timeout=15)

        if not dlg:
            logging.error(
                f"Nao foi possivel abrir dialogo '{titulo_dialogo}'. "
                "Verifique se o PVA esta na tela principal (sem escrituracao aberta)."
            )
            return False

        # 4. Seleciona todos e confirma
        self._confirmar_dialogo_batch(dlg)

        # 5. Aguarda conclusao
        timeout = self.cfg.get(timeout_chave, timeout_padrao)
        logging.info(f"Aguardando conclusao de '{titulo_dialogo}' ({timeout}s)")
        time.sleep(timeout)

        # 6. Fecha popups de resultado (Informacao, Sucesso, Erro, etc.)
        for _ in range(8):
            if not _fechar_popups():
                break
            time.sleep(1.5)

        return True

    # -- Operacoes batch (chamadas apos o usuario importar manualmente) --------

    def batch_verificar_pendencias(self) -> bool:
        """Verificar Pendencias em lote: Ctrl+V -> seleciona tudo -> OK.
        Toolbar: botao 3 (indice 0-based) — icone checkmark verde.
        Busca por 'Verificar' (sem acento) para evitar mismatch de encoding.
        """
        return self._batch_operacao(
            shortcut_keys=["ctrl", "v"],
            titulo_dialogo="Verificar",
            timeout_chave="aguardar_validacao_segundos",
            timeout_padrao=600,
            toolbar_btn_index=3,
        )

    def batch_gerar_arquivo(self) -> bool:
        """Gerar Arquivo em lote: Ctrl+G -> seleciona tudo -> OK.
        Toolbar: botao 4 — icone seta verde com documento.
        """
        return self._batch_operacao(
            shortcut_keys=["ctrl", "g"],
            titulo_dialogo="Gerar",
            timeout_chave="aguardar_geracao_segundos",
            timeout_padrao=900,
            toolbar_btn_index=4,
        )

    def batch_assinar(self) -> bool:
        """Assinar em lote: Ctrl+S -> seleciona tudo -> OK.
        Toolbar: botao 5 — icone caneta/edit.
        """
        return self._batch_operacao(
            shortcut_keys=["ctrl", "s"],
            titulo_dialogo="Assinar",
            timeout_chave="aguardar_assinatura_segundos",
            timeout_padrao=300,
            toolbar_btn_index=5,
        )

    def batch_transmitir(self) -> bool:
        """Transmitir em lote: Ctrl+T -> seleciona tudo -> OK.
        Toolbar: botao 7 — icone globo.
        """
        return self._batch_operacao(
            shortcut_keys=["ctrl", "t"],
            titulo_dialogo="Transmit",
            timeout_chave="aguardar_transmissao_segundos",
            timeout_padrao=900,
            toolbar_btn_index=7,
        )

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
            logging.error("Falha 