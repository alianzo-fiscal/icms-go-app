# coding: utf-8
"""
certidoes_rfb.py — Emissao em lote CND Federal (Receita Federal) via Playwright.

URL: https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir

Fluxo:
  1. Preenche CNPJ e clica Consultar
  2. Se aparecer "Emissao de nova certidao" -> clica -> captura PDF
  3. Se ja existir certidao valida -> clica "Consulta / 2a via" -> captura PDF

Uso:
  python certidoes_rfb.py --empresa EDN --limite 5
  python certidoes_rfb.py --empresa EDN --headless
  python certidoes_rfb.py --empresa EDN --apenas 20758851000105
"""
import sys
import time
import argparse
import json
from pathlib import Path
from datetime import date

URL_RFB = "https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir"
DELAY_ENTRE = 5

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


def _so_numeros(s):
    return "".join(c for c in str(s) if c.isdigit())


def _pagina_tem_certidao(pg, patterns=("certidao", "certidão", ".pdf")):
    try:
        url = (pg.url or "").lower()
        return not pg.is_closed() and any(p in url for p in patterns)
    except Exception:
        return False


def emitir_cnpj(page, context, cnpj_num, output_path, debug=False):
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    novas_paginas = []

    def _on_new_page(np):
        novas_paginas.append(np)

    context.on("page", _on_new_page)

    try:
        # ---- 1. Navega para o formulário ----
        page.goto(URL_RFB, timeout=30000, wait_until="networkidle")
        time.sleep(1)

        # ---- 2. Preenche CNPJ ----
        campo = page.query_selector('#NI')
        if not campo:
            resultado["msg"] = "Campo CNPJ (#NI) nao encontrado"
            return resultado
        campo.click()
        campo.fill("")
        campo.type(cnpj_num, delay=30)
        time.sleep(0.3)

        # ---- 3. Clica Consultar ----
        page.click('#validar', timeout=10000)
        time.sleep(2)

        # ---- 4. Analisa resultado ----
        texto = ""
        try:
            el = page.query_selector('#rfb-main-container > div')
            texto = el.inner_text() if el else ""
        except Exception:
            pass

        if debug:
            print(f"  Texto resposta: {texto[:200]!r}")

        # Erros conhecidos
        if "não foi encontrada" in texto:
            resultado["msg"] = "CNPJ nao encontrado"
            return resultado
        if "são insuficientes" in texto:
            resultado["msg"] = "Certidao Positiva (debitos)"
            return resultado

        # ---- 5. Clica no link de emissão/2a via ----
        link_emitir = (
            page.query_selector('a:has-text("Emissão de nova certidão")')
            or page.query_selector('a:has-text("Emissao de nova certidao")')
            or page.query_selector('a:has-text("Consulta de certidão e emissão de 2ª via")')
            or page.query_selector('a:has-text("Consulta de certidao")')
            or page.query_selector('a[href*="certidao"]')
        )

        if not link_emitir:
            # Pode estar na 2ª tela após emissao bem sucedida
            link_emitir = (
                page.query_selector('a:has-text("2ª via")')
                or page.query_selector('a:has-text("2a via")')
            )

        if not link_emitir:
            resultado["msg"] = f"Link de emissao nao encontrado. Texto={texto[:120]}"
            return resultado

        link_emitir.click()
        time.sleep(2)

        # ---- 6. Aguarda nova página com a certidão ----
        alvo = None
        for _ in range(30):
            if _pagina_tem_certidao(page):
                alvo = page
                break
            for np in novas_paginas:
                if _pagina_tem_certidao(np):
                    alvo = np
                    break
            if alvo:
                break
            # Também verifica se a página principal navegou para certidao
            try:
                cur = page.url.lower()
                if "certidao" in cur or ".pdf" in cur:
                    alvo = page
                    break
            except Exception:
                pass
            time.sleep(1)

        if alvo is None:
            # Se não encontrou por URL, tenta pegar qualquer nova página aberta
            if novas_paginas:
                alvo = novas_paginas[-1]
                try:
                    alvo.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass
            else:
                resultado["msg"] = "Pagina da certidao nao encontrada em 30s"
                return resultado

        try:
            alvo.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        if debug:
            print(f"  Certidao URL: {alvo.url}")

        # ---- 7. Salva como PDF ----
        alvo.pdf(
            path=str(output_path),
            format="A4",
            print_background=True,
        )

        try:
            if alvo != page and not alvo.is_closed():
                alvo.close()
        except Exception:
            pass

        resultado["status"] = "ok"
        resultado["arquivo"] = str(output_path)
        resultado["msg"] = "OK"

    except Exception as e:
        resultado["msg"] = str(e)
    finally:
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
    ap.add_argument("--debug",    action="store_true")
    args = ap.parse_args()

    empresa = EMPRESAS.get(args.empresa.upper())
    if not empresa:
        print(f"Empresa '{args.empresa}' nao encontrada.")
        sys.exit(1)

    base_out = Path(args.output) if args.output else Path(__file__).parent.parent / "certidoes_output"
    pasta    = base_out / "RFB" / str(date.today())
    pasta.mkdir(parents=True, exist_ok=True)
    log_path = pasta / "log.json"

    lista = empresa["cnpjs"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]
    if args.limite > 0:
        lista = lista[:args.limite]

    print(f"\n{'='*60}")
    print(f"  CND Federal (RFB) — {empresa['nome']}")
    print(f"  Total: {len(lista)} CNPJs  |  Saida: {pasta}")
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
            pdf_path = pasta / f"{tag}_{cnpj_num}.pdf"

            if pdf_path.exists():
                print(f"[{i:02d}/{len(lista)}] {tag} — ja existe, pulando")
                resultados.append({"cnpj": cnpj_num, "status": "pulado", "msg": "ja existia"})
                continue

            print(f"[{i:02d}/{len(lista)}] {tag} — {cnpj_num}", end=" ... ", flush=True)
            if args.debug:
                print()

            res = emitir_cnpj(pg, context, cnpj_num, pdf_path, debug=args.debug)
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
