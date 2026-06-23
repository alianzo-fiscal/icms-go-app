"""
Fase 2: gerar + assinar + transmitir escrituracoes ja validadas no PVA.
Pressupoe que a Fase 1 foi executada e o PVA esta aberto com as escrituracoes no banco.
Executar: python pva_fase2.py
"""
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

from fase1_lote import _carregar_config, _ler_resultados, _salvar_resultado
from pva_automacao import PVAAutomacao


def main():
    cfg      = _carregar_config()
    log_json = Path(cfg["log_validacao"])

    logging.basicConfig(
        filename=cfg["log_arquivo"],
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        encoding="utf-8",
    )
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))

    resultados = _ler_resultados(log_json)
    pendentes  = [r for r in resultados if r.get("status") == "OK" and not r.get("fase2_ok")]

    if not pendentes:
        print("Nenhum arquivo pendente para Fase 2.")
        return

    sep = "=" * 60
    print(f"\n{sep}")
    print(f"  Fase 2 — {len(pendentes)} arquivo(s) para gerar/assinar/transmitir")
    print(sep)
    for i, r in enumerate(pendentes):
        print(f"  [{i}] {r['arquivo']}")

    print("\nCertifique-se de que o PVA esta aberto com as escrituracoes")
    print("validadas na Fase 1. NAO interaja com o computador durante o processo.")
    conf = input("\nDigite SIM para continuar: ").strip().upper()
    if conf != "SIM":
        print("Cancelado.")
        return

    pva = PVAAutomacao(cfg)

    for i, r in enumerate(pendentes):
        print(f"\n[FASE2] {r['arquivo']} (posicao {i} no PVA)")
        ok = False
        try:
            ok = pva.fase2_processar(index=i)
        except Exception as e:
            logging.error(f"Excecao na Fase 2 em {r['arquivo']}: {e}")

        r["fase2_ok"]        = ok
        r["fase2_timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _salvar_resultado(log_json, resultados)
        print(f"  -> {'TRANSMITIDO' if ok else 'ERRO'}")

    try:
        pva.fechar_pva()
    except Exception:
        pass

    print(f"\n{sep}")
    ok_count = sum(1 for r in resultados if r.get("fase2_ok"))
    print(f"  Fase 2 concluida: {ok_count} transmitido(s)")
    print(sep)


if __name__ == "__main__":
    main()
