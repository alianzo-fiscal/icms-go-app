# coding: utf-8
"""
certidoes_sefaz_go.py — Emissão em lote da Certidão de Débitos SEFAZ-GO.

Fluxo correto:
  1. Navega para sefaz.go.gov.br/certidao/emissao/
  2. Seleciona radio CNPJ
  3. Preenche CNPJ
  4. Marca Espolio = Nao
  5. Clica Emitir -> abre popup com confirmacao
  6. Na popup, clica Sim -> dispara download do arquivo .asp
  7. Salva o arquivo na pasta de saida

Uso:
  python certidoes_sefaz_go.py --empresa EDN
  python certidoes_sefaz_go.py --empresa EDN --headless
  python certidoes_sefaz_go.py --empresa EDN --limite 3
  python certidoes_sefaz_go.py --empresa EDN --apenas 20758851000105
"""
import sys
import time
import argparse
import json
from pathlib import Path
from datetime import date

EMPRESAS = {
    "EDN": {
        "nome": "EDN Utilidades Domesticas",
        "cnpjs": [
            {"tag": "F003_URUACU",   "cnpj": "20758851004100"},
            {"tag": "F004_APARECID", "cnpj": "20758851004283"},
            {"tag": "F005_ITUMBIAR", "cnpj": "20758851004364"},
            {"tag": "F007_TRINDADE", "cnpj": "20758851004526"},
            {"tag": "F008_GOIANIA",  "cnpj": "20758851004798"},
            {"tag": "F010_CALDAS_N", "cnpj": "20758851004607"},
            {"tag": "F011_LUZIANIA", "cnpj": "20758851005093"},
            {"tag": "F013_ITABERAI", "cnpj": "20758851005760"},
            {"tag": "F016_GOIANIA",  "cnpj": "20758851001934"},
            {"tag": "F018_GOIANIA",  "cnpj": "20758851002230"},
            {"tag": "F020_GOIANIA",  "cnpj": "20758851003473"},
            {"tag": "F024_GOIANIA",  "cnpj": "20758851001420"},
            {"tag": "F025_INHUMAS",  "cnpj": "20758851003716"},
            {"tag": "F026_ANAPOLIS", "cnpj": "20758851005689"},
            {"tag": "F027_RIO_VERD", "cnpj": "20758851005174"},
            {"tag": "F030_GOIANIRA", "cnpj": "20758851004011"},
            {"tag": "F031_GOIANIA",  "cnpj": "20758851000105"},
            {"tag": "F033_GOIANIA",  "cnpj": "20758851006065"},
            {"tag": "F038_PORANGAT", "cnpj": "20758851005840"},
            {"tag": "F048_ANAPOLIS", "cnpj": "20758851000296"},
            {"tag": "F050_GOIANIA",  "cnpj": "20758851003554"},
            {"tag": "F060_GOIANIA",  "cnpj": "20758851003635"},
            {"tag": "F071_GOIANIA",  "cnpj": "20758851000962"},
            {"tag": "F074_TRINDADE", "cnpj": "20758851001268"},
            {"tag": "F076_SENADOR_", "cnpj": "20758851000709"},
            {"tag": "F077_APARECID", "cnpj": "20758851000881"},
            {"tag": "F079_GOIANIA",  "cnpj": "20758851001004"},
            {"tag": "F080_GOIANESI", "cnpj": "20758851001187"},
            {"tag": "F082_GOIANIA",  "cnpj": "20758851002310"},
            {"tag": "F083_GOIANIA",  "cnpj": "20758851002582"},
            {"tag": "F084_GOIANIA",  "cnpj": "20758851002400"},
            {"tag": "F085_GOIANIA",  "cnpj": "20758851001349"},
            {"tag": "F086_GOIANIA",  "cnpj": "20758851001772"},
            {"tag": "F088_GOIANIA",  "cnpj": "20758851001500"},
            {"tag": "F089_GOIANIA",  "cnpj": "20758851001853"},
            {"tag": "F090_GOIANIA",  "cnpj": "20758851002159"},
            {"tag": "F091_GOIANIA",  "cnpj": "20758851003201"},
            {"tag": "F093_GOIANIA",  "cnpj": "20758851002744"},
            {"tag": "F094_GOIANIA",  "cnpj": "20758851002825"},
            {"tag": "F095_GOIANIA",  "cnpj": "20758851003040"},
            {"tag": "F096_GOIANIA",  "cnpj": "20758851002906"},
            {"tag": "F097_GOIANIA",  "cnpj": "20758851003120"},
            {"tag": "F098_GOIANIA",  "cnpj": "20758851001691"},
        ],
    }
}

URL_SEFAZ_GO = "https://www.sefaz.go.gov.br/certidao/emissao/"
DELAY_ENTRE  = 4


def _so_numeros(s):
    return "".join(c for c in str(s) if c.isdigit())


def emitir_cnpj(page, context, cnpj_num, output_path):
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    try:
        page.goto(URL_SEFAZ_GO, timeout=20000, wait_until="domcontentloaded")
        time.sleep(2)

        # Radio CNPJ
        page.click('input[name="Certidao.TipoDocumento"][value="2"]', timeout=8000)
        time.sleep(0.4)

        # Espolio = Nao
        page.click('#Certidao\\.EspolioN', timeout=5000)

        # Preenche CNPJ
        campo = page.query_selector('input[id="Certidao.NumeroDocumentoCNPJ"]')
        if not campo:
            resultado["msg"] = "Campo CNPJ nao encontrado"
            return resultado
        campo.click()
        campo.fill("")
        campo.type(cnpj_num, delay=30)
        time.sleep(0.4)

        # Clica Emitir -> abre popup de confirmacao
        with context.expect_page() as popup_info:
            page.click('input[type="submit"][value="Emitir"]', timeout=5000)

        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        time.sleep(2)

        # Localiza botao Sim no popup
        btn_sim = popup.query_selector('#Certidao\\.ConfirmaNomeContribuinteSim')
        if not btn_sim:
            btn_sim = popup.query_selector('input[value="Sim"]')

        if not btn_sim or not btn_sim.is_visible():
            resultado["msg"] = "Botao Sim nao encontrado no popup"
            popup.close()
            return resultado

        # Clica Sim -> dispara download do arquivo .asp
        with popup.expect_download(timeout=30000) as dl_info:
            btn_sim.click()

        download = dl_info.value
        download.save_as(str(output_path))
        popup.close()

        resultado["status"] = "ok"
        resultado["arquivo"] = str(output_path)
        resultado["msg"] = "OK"

    except Exception as e:
        resultado["msg"] = str(e)

    return resultado


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--empresa",  default="EDN")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--output",   default="")
    ap.add_argument("--apenas",   default="")
    ap.add_argument("--limite",   type=int, default=0)
    args = ap.parse_args()

    empresa = EMPRESAS.get(args.empresa.upper())
    if not empresa:
        print(f"Empresa '{args.empresa}' nao encontrada.")
        sys.exit(1)

    base_out = Path(args.output) if args.output else Path(__file__).parent.parent / "certidoes_output"
    pasta    = base_out / "SEFAZ-GO" / str(date.today())
    pasta.mkdir(parents=True, exist_ok=True)
    log_path = pasta / "log.json"

    lista = empresa["cnpjs"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]
    if args.limite > 0:
        lista = lista[:args.limite]

    print(f"\n{'='*60}")
    print(f"  SEFAZ-GO — {empresa['nome']}")
    print(f"  Total: {len(lista)} CNPJs  |  Saida: {pasta}")
    print(f"{'='*60}\n")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright nao instalado. Execute: pip install playwright")
        sys.exit(1)

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=50)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        for i, item in enumerate(lista, 1):
            cnpj_num = _so_numeros(item["cnpj"])
            tag      = item["tag"]
            asp_path = pasta / f"{tag}_{cnpj_num}.asp"

            if asp_path.exists():
                print(f"[{i:02d}/{len(lista)}] {tag} — ja existe, pulando")
                resultados.append({"cnpj": cnpj_num, "status": "pulado", "arquivo": "", "msg": "ja existia"})
                continue

            print(f"[{i:02d}/{len(lista)}] {tag} — {cnpj_num}", end=" ... ", flush=True)

            res = emitir_cnpj(page, context, cnpj_num, asp_path)
            resultados.append(res)

            print("OK" if res["status"] == "ok" else f"ERRO: {res['msg']}")

            if i < len(lista):
                time.sleep(DELAY_ENTRE)

        browser.close()

    log_path.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")
    ok  = sum(1 for r in resultados if r["status"] == "ok")
    err = sum(1 for r in resultados if r["status"] == "erro")
    print(f"\n{'='*60}")
    print(f"  Concluido: {ok} OK  |  {err} erros  |  log: {log_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
