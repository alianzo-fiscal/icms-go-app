# coding: utf-8
"""
certidoes_sefaz_go.py — Emissao em lote SEFAZ-GO via Playwright.

Uso:
  python certidoes_sefaz_go.py --empresa EDN --limite 3
  python certidoes_sefaz_go.py --empresa EDN --headless
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


def emitir_cnpj(page, context, browser, cnpj_num, output_path, debug=False):
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}

    # ---- coleta unificada ----
    all_downloads = []      # objetos Download do Playwright
    attachment_bytes = []   # bytes de respostas com Content-Disposition: attachment

    def _capture_dl(dl):
        if debug:
            print(f"    [DL] url={dl.url}")
        all_downloads.append(dl)

    def _capture_resp(label, resp):
        try:
            cd = resp.headers.get("content-disposition", "")
            ct = resp.headers.get("content-type", "")
            if debug:
                print(f"    [{label}] {resp.status} ct={ct[:50]} cd={cd[:50]} url={resp.url[:60]}")
            # SEFAZ-GO serve certidao.asp sem Content-Type/Content-Disposition
            is_certidao = ("sefaz.go.gov.br" in resp.url and
                           "certidao.asp" in resp.url and
                           resp.status == 200)
            if is_certidao or "attachment" in cd.lower() or "application/octet-stream" in ct.lower():
                attachment_bytes.append(resp.body())
        except Exception:
            pass

    def _on_new_page(new_pg):
        if debug:
            print(f"    [NEW_PAGE] {new_pg.url}")
        new_pg.on("download", _capture_dl)
        new_pg.on("response", lambda r: _capture_resp("new_page", r))

    # registra para TODAS as páginas novas criadas daqui em diante
    context.on("page", _on_new_page)

    try:
        page.goto(URL_SEFAZ_GO, timeout=20000, wait_until="domcontentloaded")
        time.sleep(2)

        # Seleciona CNPJ
        page.click('input[name="Certidao.TipoDocumento"][value="2"]', timeout=8000)
        time.sleep(0.4)

        # Espólio = Não
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

        # Emitir → abre popup de confirmação
        with context.expect_page() as popup_info:
            page.click('input[type="submit"][value="Emitir"]', timeout=5000)

        popup = popup_info.value
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        time.sleep(1)

        # Registra listeners no popup
        popup.on("download", _capture_dl)
        popup.on("response", lambda r: _capture_resp("popup", r))
        # Também no main page (pode ser que o download volte para ela)
        page.on("download", _capture_dl)
        page.on("response", lambda r: _capture_resp("main", r))

        if debug:
            print(f"  Popup URL: {popup.url}")

        # Localiza botão Sim
        btn_sim = popup.query_selector('#Certidao\\.ConfirmaNomeContribuinteSim')
        if not btn_sim:
            btn_sim = popup.query_selector('input[value="Sim"]')

        if not btn_sim or not btn_sim.is_visible():
            resultado["msg"] = "Botao Sim nao encontrado"
            popup.close()
            return resultado

        # Clica Sim — pode abrir uma 3ª página OU triggerar download direto
        btn_sim.click()

        # Aguarda até 30s por qualquer captura
        for _ in range(30):
            if all_downloads or attachment_bytes:
                break
            time.sleep(1)

        # Tenta fechar o popup se ainda estiver aberto
        try:
            if not popup.is_closed():
                popup.close()
        except Exception:
            pass

        if all_downloads:
            if debug:
                print(f"  Salvando via Download object")
            all_downloads[0].save_as(str(output_path))
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (download event)"
        elif attachment_bytes:
            if debug:
                print(f"  Salvando via response body")
            output_path.write_bytes(attachment_bytes[0])
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (response body)"
        else:
            resultado["msg"] = "Arquivo nao recebido em 30s"

    except Exception as e:
        resultado["msg"] = str(e)
    finally:
        # Remove listener de novas páginas para não acumular entre CNPJs
        try:
            context.remove_listener("page", _on_new_page)
        except Exception:
            pass

    return resultado


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--empresa",  default="EDN")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--output",   default="")
    ap.add_argument("--apenas",   default="")
    ap.add_argument("--limite",   type=int, default=0)
    ap.add_argument("--debug",    action="store_true",
                    help="Exibe todas as respostas/downloads para diagnostico")
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
    if args.debug:
        print(f"  Modo DEBUG ativado")
    print(f"{'='*60}\n")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright nao instalado.")
        sys.exit(1)

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=50)
        context = browser.new_context(accept_downloads=True)
        pg      = context.new_page()

        for i, item in enumerate(lista, 1):
            cnpj_num = _so_numeros(item["cnpj"])
            tag      = item["tag"]
            asp_path = pasta / f"{tag}_{cnpj_num}.asp"

            if asp_path.exists():
                print(f"[{i:02d}/{len(lista)}] {tag} — ja existe, pulando")
                resultados.append({"cnpj": cnpj_num, "status": "pulado", "msg": "ja existia"})
                continue

            print(f"[{i:02d}/{len(lista)}] {tag} — {cnpj_num}", end=" ... ", flush=True)
            if args.debug:
                print()

            res = emitir_cnpj(pg, context, browser, cnpj_num, asp_path, debug=args.debug)
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
