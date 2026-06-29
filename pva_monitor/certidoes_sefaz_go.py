# coding: utf-8
"""
certidoes_sefaz_go.py — Emissão em lote da Certidão de Débitos SEFAZ-GO.

Para cada CNPJ da empresa:
  1. Acessa sefaz.go.gov.br/certidao/emissao/
  2. Seleciona tipo CNPJ
  3. Preenche o número
  4. Clica em Emitir
  5. Salva o PDF na pasta de saída

Uso:
  python certidoes_sefaz_go.py --empresa EDN
  python certidoes_sefaz_go.py --empresa EDN --headless
  python certidoes_sefaz_go.py --empresa EDN --output C:\\certidoes

Saída padrão:
  certidoes_output/SEFAZ-GO/YYYY-MM-DD/<filial>_<cnpj>.pdf
"""
import sys
import time
import argparse
import json
from pathlib import Path
from datetime import date

# ─────────────────────────── CNPJs EDN ───────────────────────────────────────
EMPRESAS = {
    "EDN": {
        "nome": "EDN Utilidades Domésticas",
        "cnpjs": [
            {"tag": "F003_URUACU", "cnpj": "20758851004100", "ie": "20.013.121-4", "im": "993055256", "municipio": "URUAÇU"},
            {"tag": "F004_APARECID", "cnpj": "20758851004283", "ie": "20.030.770-3", "im": "3110105706", "municipio": "Aparecida de GOIÂNIA"},
            {"tag": "F005_ITUMBIAR", "cnpj": "20758851004364", "ie": "20.036.783-8", "im": "610403", "municipio": "ITUMBIARA"},
            {"tag": "F007_TRINDADE", "cnpj": "20758851004526", "ie": "20.086.156-5", "im": "16230931", "municipio": "TRINDADE"},
            {"tag": "F008_GOIANIA", "cnpj": "20758851004798", "ie": "20.099.450-6", "im": "6415024", "municipio": "GOIÂNIA"},
            {"tag": "F010_CALDAS_N", "cnpj": "20758851004607", "ie": "20.092.115-0", "im": "66288", "municipio": "CALDAS NOVAS"},
            {"tag": "F011_LUZIANIA", "cnpj": "20758851005093", "ie": "20.116.302-0", "im": "2000009306", "municipio": "LUZIANIA"},
            {"tag": "F013_ITABERAI", "cnpj": "20758851005760", "ie": "20.161.077-9", "im": "10213", "municipio": "ITABERAI"},
            {"tag": "F016_GOIANIA", "cnpj": "20758851001934", "ie": "10.831.477-4", "im": "73320", "municipio": "GOIÂNIA"},
            {"tag": "F018_GOIANIA", "cnpj": "20758851002230", "ie": "10.827.190-0", "im": "5298741", "municipio": "GOIÂNIA"},
            {"tag": "F020_GOIANIA", "cnpj": "20758851003473", "ie": "10.899.551-8", "im": "5604291", "municipio": "GOIÂNIA"},
            {"tag": "F024_GOIANIA", "cnpj": "20758851001420", "ie": "10.827.908-1", "im": "5377986", "municipio": "GOIÂNIA"},
            {"tag": "F025_INHUMAS", "cnpj": "20758851003716", "ie": "10.940.721-0", "im": "14710", "municipio": "INHUMAS"},
            {"tag": "F026_ANAPOLIS", "cnpj": "20758851005689", "ie": "20.161.306-9", "im": "122475", "municipio": "ANAPOLIS"},
            {"tag": "F027_RIO_VERD", "cnpj": "20758851005174", "ie": "20.126.648-2", "im": "77933", "municipio": "RIO VERDE"},
            {"tag": "F030_GOIANIRA", "cnpj": "20758851004011", "ie": "10.987.544-3", "im": "11687", "municipio": "Goianira"},
            {"tag": "F031_GOIANIA", "cnpj": "20758851000105", "ie": "10.607.511-0", "im": "3834344", "municipio": "GOIÂNIA"},
            {"tag": "F033_GOIANIA", "cnpj": "20758851006065", "ie": "20.293.165-0", "im": "", "municipio": "GOIÂNIA"},
            {"tag": "F038_PORANGAT", "cnpj": "20758851005840", "ie": "20.171.456-6", "im": "651008687", "municipio": "PORANGATU"},
            {"tag": "F048_ANAPOLIS", "cnpj": "20758851000296", "ie": "10.724.310-5", "im": "86409", "municipio": "ANAPOLIS"},
            {"tag": "F050_GOIANIA", "cnpj": "20758851003554", "ie": "10.901.034-5", "im": "3110082781", "municipio": "GOIÂNIA"},
            {"tag": "F060_GOIANIA", "cnpj": "20758851003635", "ie": "10.932.688-1", "im": "5709709", "municipio": "GOIÂNIA"},
            {"tag": "F071_GOIANIA", "cnpj": "20758851000962", "ie": "10.811.073-7", "im": "5197929", "municipio": "GOIÂNIA"},
            {"tag": "F074_TRINDADE", "cnpj": "20758851001268", "ie": "10.873.370-0", "im": "16220835", "municipio": "TRINDADE"},
            {"tag": "F076_SENADOR_", "cnpj": "20758851000709", "ie": "10.806.526-0", "im": "30009692", "municipio": "SENADOR CANEDO"},
            {"tag": "F077_APARECID", "cnpj": "20758851000881", "ie": "10.807.111-1", "im": "3110058005", "municipio": "Aparecida de Gyn"},
            {"tag": "F079_GOIANIA", "cnpj": "20758851001004", "ie": "10.814.950-1", "im": "202631701", "municipio": "GOIÂNIA"},
            {"tag": "F080_GOIANESI", "cnpj": "20758851001187", "ie": "10.819.372-1", "im": "356949", "municipio": "Goianesia"},
            {"tag": "F082_GOIANIA", "cnpj": "20758851002310", "ie": "10.827.170-6", "im": "5372798", "municipio": "GOIÂNIA"},
            {"tag": "F083_GOIANIA", "cnpj": "20758851002582", "ie": "10.833.248-9", "im": "5325625", "municipio": "GOIÂNIA"},
            {"tag": "F084_GOIANIA", "cnpj": "20758851002400", "ie": "10.832.158-4", "im": "3110061704", "municipio": "GOIÂNIA"},
            {"tag": "F085_GOIANIA", "cnpj": "20758851001349", "ie": "10.828.382-8", "im": "5377927", "municipio": "GOIÂNIA"},
            {"tag": "F086_GOIANIA", "cnpj": "20758851001772", "ie": "10.832.210-6", "im": "3110061722", "municipio": "GOIÂNIA"},
            {"tag": "F088_GOIANIA", "cnpj": "20758851001500", "ie": "10.832.142-8", "im": "3110061712", "municipio": "GOIÂNIA"},
            {"tag": "F089_GOIANIA", "cnpj": "20758851001853", "ie": "10.828.043-8", "im": "5298733", "municipio": "GOIÂNIA"},
            {"tag": "F090_GOIANIA", "cnpj": "20758851002159", "ie": "10.832.205-0", "im": "3110061717", "municipio": "GOIÂNIA"},
            {"tag": "F091_GOIANIA", "cnpj": "20758851003201", "ie": "10.854.569-5", "im": "5430275", "municipio": "GOIÂNIA"},
            {"tag": "F093_GOIANIA", "cnpj": "20758851002744", "ie": "10.857.722-8", "im": "5430283", "municipio": "GOIÂNIA"},
            {"tag": "F094_GOIANIA", "cnpj": "20758851002825", "ie": "10.858.991-9", "im": "5430291", "municipio": "GOIÂNIA"},
            {"tag": "F095_GOIANIA", "cnpj": "20758851003040", "ie": "10.859.183-2", "im": "25174", "municipio": "GOIÂNIA"},
            {"tag": "F096_GOIANIA", "cnpj": "20758851002906", "ie": "10.855.552-6", "im": "3110069838", "municipio": "GOIÂNIA"},
            {"tag": "F097_GOIANIA", "cnpj": "20758851003120", "ie": "10.854.061-8", "im": "5430305", "municipio": "GOIÂNIA"},
            {"tag": "F098_GOIANIA", "cnpj": "20758851001691", "ie": "10.827.173-0", "im": "5298725", "municipio": "GOIÂNIA"},
        ],
    }
}



URL_SEFAZ = "https://www.sefaz.go.gov.br/certidao/emissao/"
DELAY_ENTRE = 4   # segundos entre cada CNPJ


def _so_numeros(s: str) -> str:
    return "".join(c for c in s if c.isdigit())


def emitir_cnpj(page, context, cnpj_num: str, output_path: Path) -> dict:
    """Emite a certidão para um CNPJ e salva o PDF. Retorna dict com status."""
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    try:
        page.goto(URL_SEFAZ, timeout=20000, wait_until="domcontentloaded")
        time.sleep(2)

        # 1. Seleciona radio CNPJ — value="2" no formulário real
        page.click('input[name="Certidao.TipoDocumento"][value="2"]', timeout=8000)
        time.sleep(0.5)

        # 2. Preenche campo CNPJ (id="Certidao.NumeroDocumentoCNPJ")
        campo = page.query_selector('input[id="Certidao.NumeroDocumentoCNPJ"]')
        if not campo:
            resultado["msg"] = "Campo Certidao.NumeroDocumentoCNPJ não encontrado"
            return resultado

        campo.click()
        campo.fill("")
        campo.type(cnpj_num, delay=30)
        time.sleep(0.5)

        # 3. Clica em Emitir — o site abre popup com confirmação
        with context.expect_page() as popup_info:
            page.click('input[type="submit"][value="Emitir"]', timeout=5000)

        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        time.sleep(2)

        # 4. Popup pode ser a confirmação "Confirma o Nome..." ou já a certidão
        btn_sim = popup.query_selector('input[value="Sim"], button:has-text("Sim")')
        if btn_sim and btn_sim.is_visible():
            # Há uma segunda etapa de confirmação — pode abrir outro popup
            try:
                with context.expect_page() as cert_info:
                    btn_sim.click()
                cert_page = cert_info.value
                cert_page.wait_for_load_state("load", timeout=20000)
                time.sleep(3)
                cert_page.pdf(path=str(output_path), print_background=True)
                cert_page.close()
            except Exception:
                # Clicou Sim mas certidão ficou no mesmo popup
                time.sleep(4)
                popup.pdf(path=str(output_path), print_background=True)
            popup.close()
        else:
            # Popup já é a certidão (sem etapa de confirmação)
            time.sleep(3)
            popup.pdf(path=str(output_path), print_background=True)
            popup.close()

        resultado["status"] = "ok"
        resultado["arquivo"] = str(output_path)
        resultado["msg"] = "PDF salvo"

    except Exception as e:
        resultado["msg"] = str(e)

    return resultado


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--empresa",  default="EDN", help="Empresa (ex: EDN)")
    ap.add_argument("--headless", action="store_true", help="Rodar sem janela (modo silencioso)")
    ap.add_argument("--output",   default="", help="Pasta de saída (padrão: certidoes_output)")
    ap.add_argument("--apenas",   default="", help="Rodar apenas este CNPJ (14 dígitos)")
    ap.add_argument("--limite",   type=int, default=0, help="Limitar aos primeiros N CNPJs")
    args = ap.parse_args()

    empresa = EMPRESAS.get(args.empresa.upper())
    if not empresa:
        print(f"Empresa '{args.empresa}' não encontrada. Disponíveis: {list(EMPRESAS.keys())}")
        sys.exit(1)

    # Pasta de saída
    base_out = Path(args.output) if args.output else Path(__file__).parent.parent / "certidoes_output"
    pasta = base_out / "SEFAZ-GO" / str(date.today())
    pasta.mkdir(parents=True, exist_ok=True)
    log_path = pasta / "log.json"

    # Filtra CNPJs se --apenas foi passado
    lista = empresa["cnpjs"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]
        if not lista:
            print(f"CNPJ '{args.apenas}' não encontrado na empresa {args.empresa}.")
            sys.exit(1)
    if args.limite > 0:
        lista = lista[:args.limite]

    print(f"\n{'='*60}")
    print(f"  SEFAZ-GO — Emissão em lote — {empresa['nome']}")
    print(f"  Total: {len(lista)} CNPJs")
    print(f"  Saída: {pasta}")
    print(f"{'='*60}\n")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Playwright não instalado. Execute: pip install playwright && playwright install chromium")
        sys.exit(1)

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=30)
        context = browser.new_context()
        page = context.new_page()

        for i, item in enumerate(lista, 1):
            cnpj_num = _so_numeros(item["cnpj"])
            tag = item["tag"]
            pdf_path = pasta / f"{tag}_{cnpj_num}.pdf"

            print(f"[{i:02d}/{len(lista)}] {tag} — {cnpj_num}", end=" ... ", flush=True)

            if pdf_path.exists():
                print("já existe, pulando")
                resultados.append({"cnpj": cnpj_num, "status": "pulado", "arquivo": str(pdf_path), "msg": "já existia"})
                continue

            res = emitir_cnpj(page, context, cnpj_num, pdf_path)
            resultados.append(res)

            if res["status"] == "ok":
                print("✅ OK")
            else:
                print(f"❌ {res['msg']}")

            # Pausa entre requisições
            if i < len(lista):
                time.sleep(DELAY_ENTRE)

        browser.close()

    # Salva log
    log_path.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")

    ok  = sum(1 for r in resultados if r["status"] == "ok")
    err = sum(1 for r in resultados if r["status"] == "erro")
    print(f"\n{'='*60}")
    print(f"  Concluído: {ok} OK  |  {err} erros  |  log: {log_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
