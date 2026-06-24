"""
PVA Batch — Validar, Gerar, Assinar e Transmitir em lote.

PRE-REQUISITO: o usuario deve ter:
  1. Aberto o PVA manualmente
  2. Importado os arquivos TXT via botao Importar (icone + no PVA)

Este script executa as 4 etapas em sequencia sem interacao manual:
  [1/4] Verificar Pendencias
  [2/4] Gerar Arquivo de Escrituracao para Entrega
  [3/4] Assinar
  [4/4] Transmitir

Configuracao em config.json (mesmos timeouts usados pelo PVAAutomacao):
  aguardar_validacao_segundos  (padrao: 600)
  aguardar_geracao_segundos    (padrao: 900)
  aguardar_assinatura_segundos (padrao: 300)
  aguardar_transmissao_segundos(padrao: 900)
"""
import json
import logging
import sys
from pathlib import Path

_HOME     = Path.home()
_MONITOR  = _HOME / "Claude" / "Projects" / "TXT_SPED_MONITOR"
_LOG_TXT  = _MONITOR / "pva_monitor.log"
_CFG_FILE = Path(__file__).parent / "config.json"


def _carregar_config() -> dict:
    cfg = {
        "pasta_monitorada"               : str(_MONITOR),
        "pva_executavel"                 : r"C:\Arquivos de Programas RFB\Programas SPED\Fiscal\SpedEFD.exe",
        "pva_titulo_janela"              : "EFD ICMS/IPI",
        "aguardar_pva_abrir_segundos"    : 40,
        "aguardar_validacao_segundos"    : 600,
        "aguardar_geracao_segundos"      : 900,
        "aguardar_assinatura_segundos"   : 300,
        "aguardar_transmissao_segundos"  : 900,
        "log_arquivo"                    : str(_LOG_TXT),
    }
    if _CFG_FILE.exists():
        dados = json.loads(_CFG_FILE.read_text(encoding="utf-8"))
        for k, v in dados.items():
            if v:
                cfg[k] = v
    return cfg


def main():
    cfg = _carregar_config()

    _MONITOR.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        filename=cfg["log_arquivo"],
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        encoding="utf-8",
    )
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))

    from pva_automacao import PVAAutomacao, _encontrar_janela

    titulo = cfg.get("pva_titulo_janela", "EFD ICMS/IPI")
    if not _encontrar_janela(titulo):
        msg = (
            "ERRO: PVA nao esta aberto.\n"
            "Abra o PVA e importe os arquivos manualmente antes de executar este script."
        )
        print(msg)
        logging.error(msg)
        sys.exit(1)

    sep = "=" * 60
    print(f"\n{sep}")
    print("  PVA Batch — Verificar / Gerar / Assinar / Transmitir")
    print(f"{sep}\n")
    print("PVA encontrado. Iniciando operacoes em lote.")
    print("NAO INTERAJA COM O COMPUTADOR DURANTE O PROCESSO.\n")

    pva = PVAAutomacao(cfg)

    etapas = [
        ("Verificar Pendencias", pva.batch_verificar_pendencias),
        ("Gerar Arquivo",        pva.batch_gerar_arquivo),
        ("Assinar",              pva.batch_assinar),
        ("Transmitir",           pva.batch_transmitir),
    ]

    for i, (nome, func) in enumerate(etapas, 1):
        print(f"[{i}/{len(etapas)}] {nome}...")
        logging.info(f"=== Iniciando: {nome} ===")
        ok = False
        try:
            ok = func()
        except Exception as exc:
            logging.error(f"Excecao em '{nome}': {exc}", exc_info=True)

        status = "OK" if ok else "ERRO"
        print(f"  -> {status}")
        logging.info(f"=== Concluido: {nome} — {status} ===")

        if not ok:
            print(f"\nFalha em: {nome}. Verifique o log: {cfg['log_arquivo']}")
            sys.exit(1)

        import time as _t
        _t.sleep(2)

    print(f"\n{sep}")
    print("  PROCESSO CONCLUIDO COM SUCESSO!")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
