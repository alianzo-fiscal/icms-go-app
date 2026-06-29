# coding: utf-8
"""
certidoes_sefaz_go.py — Emissão em lote da Certidão de Débitos SEFAZ-GO.

Fluxo correto (baseado no RPA alianzo-consulta-cnd-estadual):
  1. Navega para sefaz.go.gov.br/certidao/emissao/
  2. Seleciona radio CNPJ (value=2, id=Certidao.TipoDocumentoCNPJ)
  3. Preenche campo CNPJ (name=Certidao.NumeroDocumentoCNPJ)
  4. Marca Espólio = Não (id=Certidao.EspolioN)
  5. Clica Emitir → abre nova aba com confirmação
  6. Na nova aba, clica Sim (id=Certidao.ConfirmaNomeContribuinteSim)
  7. Aguarda o download do arquivo .asp
  8. Move/renomeia para pasta de saída

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
import unicodedata
import re
import shutil
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
            {"tag": "F010_CALDAS_N", "cnpj": "20758851004607", "ie": "20.115.832-8", "im": "84827", "municipio": "CALDAS NOVAS"},
            {"tag": "F011_LUZIANIA", "cnpj": "20758851005093", "ie": "20.119.891-5", "im": "51571", "municipio": "LUZIÂNIA"},
            {"tag": "F013_ITABERAI", "cnpj": "20758851005760", "ie": "20.159.754-7", "im": "25819", "municipio": "ITABERAÍ"},
            {"tag": "F016_GOIANIA", "cnpj": "20758851001934", "ie": "20.270.765-6", "im": "17619843", "municipio": "GOIÂNIA"},
            {"tag": "F018_GOIANIA", "cnpj": "20758851002230", "ie": "20.299.799-1", "im": "19062760", "municipio": "GOIÂNIA"},
            {"tag": "F020_GOIANIA", "cnpj": "20758851003473", "ie": "20.417.684-2", "im": "22345862", "municipio": "GOIÂNIA"},
            {"tag": "F024_GOIANIA", "cnpj": "20758851001420", "ie": "20.209.513-4", "im": "13866093", "municipio": "GOIÂNIA"},
            {"tag": "F025_INHUMAS", "cnpj": "20758851003716", "ie": "20.439.703-6", "im": "41830", "municipio": "INHUMAS"},
            {"tag": "F026_ANAPOLIS", "cnpj": "20758851005689", "ie": "20.508.279-0", "im": "78218", "municipio": "ANÁPOLIS"},
            {"tag": "F027_RIO_VERD", "cnpj": "20758851005174", "ie": "20.472.607-1", "im": "30543", "municipio": "RIO VERDE"},
            {"tag": "F030_GOIANIRA", "cnpj": "20758851004011", "ie": "20.620.521-1", "im": "10980", "municipio": "GOIANIRA"},
            {"tag": "F031_GOIANIA", "cnpj": "20758851000105", "ie": "10.607.511-0", "im": "5848016", "municipio": "GOIÂNIA"},
            {"tag": "F033_GOIANIA", "cnpj": "20758851006065", "ie": "20.597.507-8", "im": "26866965", "municipio": "GOIÂNIA"},
            {"tag": "F038_PORANGAT", "cnpj": "20758851005840", "ie": "20.528.906-6", "im": "24040", "municipio": "PORANGATU"},
            {"tag": "F048_ANAPOLIS", "cnpj": "20758851000296", "ie": "10.639.082-0", "im": "78218", "municipio": "ANÁPOLIS"},
            {"tag": "F050_GOIANIA", "cnpj": "20758851003554", "ie": "20.424.219-9", "im": "23017869", "municipio": "GOIÂNIA"},
            {"tag": "F060_GOIANIA", "cnpj": "20758851003635", "ie": "20.431.956-2", "im": "23647369", "municipio": "GOIÂNIA"},
            {"tag": "F071_GOIANIA", "cnpj": "20758851000962", "ie": "10.667.249-8", "im": "8552512", "municipio": "GOIÂNIA"},
            {"tag": "F074_TRINDADE", "cnpj": "20758851001268", "ie": "10.698.183-0", "im": "25041", "municipio": "TRINDADE"},
            {"tag": "F076_SENADOR_", "cnpj": "20758851000709", "ie": "10.660.059-3", "im": "15005", "municipio": "SENADOR CANEDO"},
            {"tag": "F077_APARECID", "cnpj": "20758851000881", "ie": "10.664.934-0", "im": "3052606910", "municipio": "Aparecida de GOIÂNIA"},
            {"tag": "F079_GOIANIA", "cnpj": "20758851001004", "ie": "10.671.025-2", "im": "8776019", "municipio": "GOIÂNIA"},
            {"tag": "F080_GOIANESI", "cnpj": "20758851001187", "ie": "10.688.116-5", "im": "28350", "municipio": "GOIANÉSIA"},
            {"tag": "F082_GOIANIA", "cnpj": "20758851002310", "ie": "20.306.744-8", "im": "19725614", "municipio": "GOIÂNIA"},
            {"tag": "F083_GOIANIA", "cnpj": "20758851002582", "ie": "20.340.344-9", "im": "20940597", "municipio": "GOIÂNIA"},
            {"tag": "F084_GOIANIA", "cnpj": "20758851002400", "ie": "20.313.636-2", "im": "20166680", "municipio": "GOIÂNIA"},
            {"tag": "F085_GOIANIA", "cnpj": "20758851001349", "ie": "10.703.793-0", "im": "12729272", "municipio": "GOIÂNIA"},
            {"tag": "F086_GOIANIA", "cnpj": "20758851001772", "ie": "20.258.988-2", "im": "16427398", "municipio": "GOIÂNIA"},
            {"tag": "F088_GOIANIA", "cnpj": "20758851001500", "ie": "20.225.406-4", "im": "14567015", "municipio": "GOIÂNIA"},
            {"tag": "F089_GOIANIA", "cnpj": "20758851001853", "ie": "20.289.375-5", "im": "17296001", "municipio": "GOIÂNIA"},
            {"tag": "F090_GOIANIA", "cnpj": "20758851002159", "ie": "20.294.625-0", "im": "18720773", "municipio": "GOIÂNIA"},
            {"tag": "F091_GOIANIA", "cnpj": "20758851003201", "ie": "20.394.629-2", "im": "21629148", "municipio": "GOIÂNIA"},
            {"tag": "F093_GOIANIA", "cnpj": "20758851002744", "ie": "20.354.282-1", "im": "21233428", "municipio": "GOIÂNIA"},
            {"tag": "F094_GOIANIA", "cnpj": "20758851002825", "ie": "20.363.157-0", "im": "21340036", "municipio": "GOIÂNIA"},
            {"tag": "F095_GOIANIA", "cnpj": "20758851003040", "ie": "20.382.327-2", "im": "21522615", "municipio": "GOIÂNIA"},
            {"tag": "F096_GOIANIA", "cnpj": "20758851002906", "ie": "20.370.997-7", "im": "21437059", "municipio": "GOIÂNIA"},
            {"tag": "F097_GOIANIA", "cnpj": "20758851003120", "ie": "20.387.455-7", "im": "21574773", "municipio": "GOIÂNIA"},
            {"tag": "F098_GOIANIA", "cnpj": "20758851001691", "ie": "20.249.038-8", "im": "15697413", "municipio": "GOIÂNIA"},
        ],
    }
}

URL_SEFAZ_GO = "https://www.sefaz.go.gov.br/certidao/emissao/"
DELAY_ENTRE   = 4


def _so_numeros(s):
    return "".join(c for c in str(s) if c.isdigit())


def _wait_download(download_dir: Path, timeout: int = 30) -> Path | None:
    """Aguarda até aparecer um arquivo na pasta de download."""
    for _ in range(timeout):
        files = [f for f in download_dir.iterdir() if f.is_file() and not f.suffix == ".crdownload"]
        if files:
            return files[0]
        time.sleep(1)
    return None


def emitir_cnpj(page, context, cnpj_num: str, download_dir: Path, output_path: Path) -> dict:
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    try:
        # Limpa pasta de download antes de cada emissão
        for f in download_dir.iterdir():
            f.unlink()

        page.goto(URL_SEFAZ_GO, timeout=20000, wait_until="domcontentloaded")
        time.sleep(2)

        # 1. Radio CNPJ
        page.click('input[name="Certidao.TipoDocumento"][value="2"]', timeout=8000)
        time.sleep(0.4)

        # 2. Espólio = Não
        page.click('#Certidao\\.EspolioN', timeout=5000)

        # 3. Preenche CNPJ
        campo = page.query_selector('input[id="Certidao.NumeroDocumentoCNPJ"]')
        if not campo:
            resultado["msg"] = "Campo CNPJ não encontrado"
            return resultado
        campo.click()
        campo.fill("")
        campo.type(cnpj_num, delay=30)
        time.sleep(0.4)

        # 4. Clica Emitir → abre nova aba com confirmação
        with context.expect_page() as popup_info:
            page.click('input[type="submit"][value="Emitir"]', timeout=5000)

        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        time.sleep(2)

        # 5. Na nova aba, clica Sim para confirmar o contribuinte
        #    → isso dispara o download do arquivo .asp
        btn_sim = popup.query_selector('#Certidao\\.ConfirmaNomeContribuinteSim')
        if not btn_sim:
            btn_sim = popup.query_selector('input[value="Sim"], button:has-text("Sim")')
        
        if btn_sim and btn_sim.is_visible():
            btn_sim.click()
            time.sleep(2)

        popup.close()

        # 6. Aguarda download na pasta configurada
        arquivo = _wait_download(download_dir, timeout=30)
        if not arquivo:
            resultado["msg"] = "Download não ocorreu em 30s"
            return resultado

        # 7. Move para pasta de saída
        shutil.move(str(arquivo), str(output_path))
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
        print(f"Empresa '{args.empresa}' não encontrada.")
        sys.exit(1)

    base_out   = Path(args.output) if args.output else Path(__file__).parent.parent / "certidoes_output"
    pasta      = base_out / "SEFAZ-GO" / str(date.today())
    pasta.mkdir(parents=True, exist_ok=True)
    download_dir = base_out / "SEFAZ-GO" / "_downloads"
    download_dir.mkdir(parents=True, exist_ok=True)
    log_path   = pasta / "log.json"

    lista = empresa["cnpjs"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]
    if args.limite > 0:
        lista = lista[:args.limite]

    print(f"\n{'='*60}")
    print(f"  SEFAZ-GO — {empresa['nome']}")
    print(f"  Total: {len(lista)} CNPJs  |  Saída: {pasta}")
    print(f"{'='*60}\n")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Playwright não instalado. Execute: pip install playwright")
        sys.exit(1)

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=args.headless,
            slow_mo=50,
            downloads_path=str(download_dir),
        )
        context = browser.new_context(
            accept_downloads=True,
        )
        # Configura download path via CDP
        context.grant_permissions([])
        page = context.new_page()

        # Configura pasta de download via CDP
        page.context.tracing  # garante inicialização
        try:
            cdp = context.new_cdp_session(page)
            cdp.send("Browser.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": str(download_dir),
                "eventsEnabled": True,
            })
        except Exception:
            pass

        for i, item in enumerate(lista, 1):
            cnpj_num  = _so_numeros(item["cnpj"])
            tag       = item["tag"]
            # Extensão .asp pois o SEFAZ-GO baixa arquivo ASP
            pdf_path  = pasta / f"{tag}_{cnpj_num}.asp"
            pdf_final = pasta / f"{tag}_{cnpj_num}.pdf"

            # Pula se já existe (qualquer formato)
            if pdf_path.exists() or pdf_final.exists():
                print(f"[{i:02d}/{len(lista)}] {tag} — já existe, pulando")
                resultados.append({"cnpj": cnpj_num, "status": "pulado", "arquivo": "", "msg": "já existia"})
                continue

            print(f"[{i:02d}/{len(lista)}] {tag} — {cnpj_num}", end=" ... ", flush=True)

            res = emitir_cnpj(page, context, cnpj_num, download_dir, pdf_path)
            resultados.append(res)

            if res["status"] == "ok":
                print("✅ OK")
            else:
                print(f"❌ {res['msg']}")

            if i < len(lista):
                time.sleep(DELAY_ENTRE)

        browser.close()

    log_path.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")
    ok  = sum(1 for r in resultados if r["status"] == "ok")
    err = sum(1 for r in resultados if r["status"] == "erro")
    print(f"\n{'='*60}")
    print(f"  Concluído: {ok} OK  |  {err} erros  |  log: {log_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
