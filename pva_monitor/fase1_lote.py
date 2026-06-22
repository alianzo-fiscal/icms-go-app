"""
Fase 1 em lote: importa e valida todos os .txt presentes na pasta monitorada.
Executar: python fase1_lote.py   ou   fase1_lote.bat
"""
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

# Caminhos derivados de Path.home() — funcionam em qualquer maquina do time
_HOME        = Path.home()
_MONITOR     = _HOME / "Claude" / "Projects" / "TXT_SPED_MONITOR"
_VALIDADOS   = _MONITOR / "Validados"
_COM_ERRO    = _MONITOR / "ComErro"
_LOG_JSON    = _MONITOR / "resultado_validacao.json"
_LOG_TXT     = _MONITOR / "pva_monitor.log"

# Configuracao PVA (config.json sobrescreve os defaults acima se existir)
_CFG_FILE = Path(__file__).parent / "config.json"

def _carregar_config() -> dict:
    cfg = {
        "pasta_monitorada"              : str(_MONITOR),
        "pasta_validados"               : str(_VALIDADOS),
        "pasta_com_erro"                : str(_COM_ERRO),
        "pasta_entrega"                 : str(_HOME / "Claude" / "Projects" / "TXT_ENTREGA"),
        "pva_executavel"                : r"C:\Arquivos de Programas RFB\Programas SPED\Fiscal\SpedEFD.exe",
        "pva_titulo_janela"             : "EFD ICMS/IPI",
        "aguardar_pva_abrir_segundos"   : 40,
        "aguardar_importacao_segundos"  : 90,
        "aguardar_validacao_segundos"   : 20,
        "aguardar_geracao_segundos"     : 30,
        "aguardar_assinatura_segundos"  : 60,
        "aguardar_transmissao_segundos" : 120,
        "log_arquivo"                   : str(_LOG_TXT),
        "log_validacao"                 : str(_LOG_JSON),
    }
    if _CFG_FILE.exists():
        dados = json.loads(_CFG_FILE.read_text(encoding="utf-8"))
        for k, v in dados.items():
            if v:  # sobrescreve apenas se nao vazio
                cfg[k] = v
    return cfg


def _ler_resultados(log_json: Path) -> list:
    if log_json.exists():
        try:
            return json.loads(log_json.read_text(encoding="utf-8"))
        except Exception:
            return []
    return []


def _salvar_resultado(log_json: Path, resultados: list):
    log_json.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    cfg = _carregar_config()

    pasta   = Path(cfg["pasta_monitorada"])
    validos = Path(cfg["pasta_validados"])
    erros   = Path(cfg["pasta_com_erro"])
    log_json = Path(cfg["log_validacao"])

    for p in [pasta, validos, erros]:
        p.mkdir(parents=True, exist_ok=True)

    logging.basicConfig(
        filename=cfg["log_arquivo"],
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        encoding="utf-8",
    )
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))

    txts = sorted(pasta.glob("*.txt"))
    if not txts:
        print("Nenhum arquivo .txt encontrado na pasta monitorada.")
        print(f"  Pasta: {pasta}")
        return

    print(f"\n{'='*60}")
    print(f"  Fase 1 em Lote — {len(txts)} arquivo(s) encontrado(s)")
    print(f"{'='*60}\n")

    from pva_automacao import PVAAutomacao
    pva = PVAAutomacao(cfg)
    resultados = _ler_resultados(log_json)
    ja_processados = {r["arquivo"] for r in resultados}

    for arq in txts:
        if arq.name in ja_processados:
            print(f"[SKIP] {arq.name} — ja processado anteriormente")
            continue

        print(f"\n[PROC] {arq.name}")
        ok = False
        try:
            ok = pva.fase1_processar(arq)
        except Exception as e:
            logging.error(f"Excecao ao processar {arq.name}: {e}")

        status = "OK" if ok else "ERRO"
        destino = validos if ok else erros
        try:
            arq.replace(destino / arq.name)  # replace() sobrescreve se ja existir
        except Exception as e:
            logging.warning(f"Nao foi possivel mover {arq.name}: {e}")

        entrada = {
            "arquivo"   : arq.name,
            "status"    : status,
            "timestamp" : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        resultados.append(entrada)
        _salvar_resultado(log_json, resultados)
        print(f"  -> {status}")

    print(f"\n{'='*60}")
    ok_count  = sum(1 for r in resultados if r.get("status") == "OK")
    err_count = sum(1 for r in resultados if r.get("status") != "OK")
    print(f"  Concluido: {ok_count} OK  |  {err_count} com erro")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
