"""
Fase 2: gerar + assinar + transmitir arquivos validados na Fase 1.
Le resultado_validacao.json e processa apenas itens com status=OK sem fase2_ok.
Requer certificado digital e-CNPJ configurado no PVA.
Executar: python pva_fase2.py   ou   fase2_assinar_transmitir.bat
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
    print(f"  Fase 2 -- {len(pendentes)} arquivo(s) para transmitir")
    print(sep)
    for r in pendentes:
        print(f"  * {r['arquivo']}")

    print("\nPRE-REQUISITO: certificado digital e-CNPJ deve estar")
    print("configurado no PVA (Configuracoes -> Certificado Digital).")
    conf = input("\nDigite SIM para continuar: ").strip().upper()
    if conf != "SIM":
        print("Cancelado.")
        return

    pva = PVAAutomacao(cfg)
    pasta_validados = Path(cfg["pasta_validados"])

    for i, r in enumerate(pendentes):
        arq = pasta_validados / r["arquivo"]
        print(f"\n[FASE2] {r['arquivo']} (posicao {i} no PVA)")
        ok = False
        try:
            ok = pva.fase2_processar(arq, index=i)
        except Exception as e:
            logging.error(f"Excecao na Fase 2 em {r['arquivo']}: {e}")

        r["fase2_ok"]        = ok
        r["fase2_timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _salvar_resultado(log_json, resultados)
        print(f"  -> {'TRANSMITIDO' if ok else 'ERRO'}")

    print(f"\n{sep}")
    ok_count = sum(1 for r in resultados if r.get("fase2_ok"))
    print(f"  Fase 2 concluida: {ok_count} transmitido(s)")
    print(sep)


if __name__ == "__main__":
    main()
