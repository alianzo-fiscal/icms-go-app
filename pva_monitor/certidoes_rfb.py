# coding: utf-8
"""
certidoes_rfb.py — Emissao CND Federal (Receita Federal) via Playwright.

IMPORTANTE: A Receita Federal so emite certidao para o CNPJ da MATRIZ.
Todos os CNPJs filiais de uma mesma empresa compartilham UMA certidao.

Novo portal (jul/2025): https://servicos.receitafederal.gov.br/servico/certidoes/

Fluxo:
  1. Navega para #/home/cnpj
  2. Preenche CNPJ da matriz e clica "Emitir Certidao"
  3a. Modal "Certidao Valida Encontrada" -> clica "Consultar Certidao"
       -> pagina /consultar -> clica "Consultar Certidao"
       -> pagina /consultar/resultado -> clica icone 2a Via
  3b. Sem modal -> certidao emitida diretamente (PDF abre/baixa)
  4. Captura PDF via nova pagina ou download

Uso:
  python certidoes_rfb.py --empresa EDN
  python certidoes_rfb.py --empresa EDN --headless
  python certidoes_rfb.py --empresa EDN --debug
"""
import sys
import time
import argparse
import json
from pathlib import Path
from datetime import date

URL_RFB = "https://servicos.receitafederal.gov.br/servico/certidoes/#/home/cnpj"
DELAY_ENTRE = 5

# Apenas CNPJ MATRIZES (branch 0001)
# Extraido de Inscricoes 1.xlsx / INSCRICAO MUNICIPAL
EMPRESAS = {
    "EDN": {
        "nome": "EDN Utilidades Domesticas",
        "matrizes": [
            {"tag": "EDN_UTILIDADES_", "cnpj": "20758851000105", "razao": "EDN UTILIDADES DOMESTICAS"},
        ],
    },
    "GRUPO": {
        "nome": "Grupo Alianzo — todas as matrizes",
        "matrizes": [
            {"tag": "AERO_PARTICIPAC", "cnpj": "53532326000126", "razao": "AERO PARTICIPACOES SPE LTDA"},
            {"tag": "AGUIA_REPRESENT", "cnpj": "39584475000108", "razao": "AGUIA REPRESENTACOES E SERVICOS LTDA"},
            {"tag": "ATACADAO_DO_LAR", "cnpj": "35917755000130", "razao": "ATACADAO DO LAR"},
            {"tag": "BIG_ESTACIONAME", "cnpj": "35279422000122", "razao": "BIG ESTACIONAMENTO"},
            {"tag": "CGV_EMPREENDIME", "cnpj": "23247575000109", "razao": "CGV EMPREENDIMENTOS"},
            {"tag": "CLUB_BIG_DECOR_", "cnpj": "58177538000156", "razao": "CLUB BIG DECOR"},
            {"tag": "COMERCIAL_DE_BR", "cnpj": "07515610000177", "razao": "COMERCIAL DE BRINQUEDOS CRISTAL"},
            {"tag": "DN_ARMAZENAMENT", "cnpj": "28221185000183", "razao": "DN ARMAZENAMENTO E TRANSPORTES"},
            {"tag": "DUARTE_E_NOGUEI", "cnpj": "33522354000155", "razao": "DUARTE E NOGUEIRA PARTICIPACOES"},
            {"tag": "EDI_INVESTIMENT", "cnpj": "33606253000162", "razao": "EDI INVESTIMENTOS"},
            {"tag": "EDN_UTILIDADES_", "cnpj": "20758851000105", "razao": "EDN UTILIDADES DOMESTICAS"},
            {"tag": "EDP_PARTICIPACO", "cnpj": "33764688000135", "razao": "EDP PARTICIPACOES"},
            {"tag": "KRISTAL_PARIS_G", "cnpj": "07877161000107", "razao": "KRISTAL PARIS GO COMERCIO DE OCULOS"},
            {"tag": "R3_SUPRIMENTOS_", "cnpj": "10641901000116", "razao": "R3 Suprimentos Corporativos"},
            {"tag": "RCI_INVESTIMENT", "cnpj": "33650317000122", "razao": "RCI INVESTIMENTOS"},
            {"tag": "RDP_PARTICIPACO", "cnpj": "33522454000181", "razao": "RDP PARTICIPACOES"},
        ],
    },
}


def _so_numeros(s):
    return "".join(c for c in str(s) if c.isdigit())


def _pagina_em(pg, *partes):
    try:
        url = (pg.url or "").lower()
        return not pg.is_closed() and any(p in url for p in partes)
    except Exception:
        return False


def emitir_cnpj(page, context, cnpj_num, output_path, debug=False):
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    novas_paginas = []
    downloads = []

    def _on_new_page(np):
        novas_paginas.append(np)
        np.on("download", lambda d: downloads.append(d))

    def _on_download(dl):
        downloads.append(dl)

    context.on("page", _on_new_page)
    page.on("download", _on_download)

    try:
        # ---- 1. Navega para o formulario CNPJ ----
        page.goto(URL_RFB, timeout=30000, wait_until="networkidle")
        time.sleep(2)

        if debug:
            print(f"  URL: {page.url}")

        # ---- 2. Preenche CNPJ ----
        campo = page.query_selector('input[placeholder="Informe o CNPJ"]')
        if not campo:
            resultado["msg"] = "Campo CNPJ nao encontrado"
            return resultado
        campo.click()
        campo.fill("")
        campo.type(cnpj_num, delay=40)
        time.sleep(0.5)

        # ---- 3. Clica "Emitir Certidao" ----
        btn_emitir = (
            page.query_selector('button[type="submit"]:has-text("Emitir Certidão")')
            or page.query_selector('button[type="submit"]')
        )
        if not btn_emitir:
            resultado["msg"] = "Botao Emitir Certidao nao encontrado"
            return resultado
        btn_emitir.click()
        time.sleep(4)

        if debug:
            print(f"  URL apos Emitir: {page.url}")

        # ---- 4. Verifica se apareceu modal de "Certidao Valida" ----
        modal_texto = page.query_selector('text="Certidão Válida Encontrada"')

        if modal_texto:
            if debug:
                print("  Modal 'Certidao Valida' detectado")

            # Clica "Consultar Certidao" no modal (botao outline/secundario)
            # Ha dois botoes: "Consultar Certidao" e "Emitir Nova Certidao"
            # Queremos o primeiro (Consultar)
            btns = page.query_selector_all('button:has-text("Consultar Certidão")')
            btn_consultar_modal = btns[0] if btns else None

            if not btn_consultar_modal:
                # Tenta texto alternativo
                btn_consultar_modal = page.query_selector('button:has-text("Consultar")')

            if not btn_consultar_modal:
                resultado["msg"] = "Botao Consultar nao encontrado no modal"
                return resultado

            btn_consultar_modal.click()
            time.sleep(3)

            if debug:
                print(f"  URL apos Consultar modal: {page.url}")

            # ---- 5. Pagina /consultar — clica "Consultar Certidao" ----
            page.wait_for_url("**/cnpj/consultar**", timeout=10000)
            time.sleep(1)

            btn_consultar_pg = page.query_selector('button:has-text("Consultar Certidão")')
            if not btn_consultar_pg:
                btn_consultar_pg = page.query_selector('button:has-text("Consultar")')

            if not btn_consultar_pg:
                resultado["msg"] = "Botao Consultar na pagina /consultar nao encontrado"
                return resultado

            btn_consultar_pg.click()
            time.sleep(4)

            if debug:
                print(f"  URL apos Consultar pag: {page.url}")

            # ---- 6. Pagina /resultado — clica icone 2a Via (1a linha) ----
            page.wait_for_url("**/consultar/resultado**", timeout=10000)
            time.sleep(1)

            # O icone de download fica na ultima coluna "2a Via" da primeira linha
            btn_2avia = (
                page.query_selector('table tbody tr:first-child td:last-child a')
                or page.query_selector('table tbody tr:first-child td:last-child button')
                or page.query_selector('tbody tr:first-child .br-button')
                or page.query_selector('tbody tr:first-child [aria-label*="Via"]')
                or page.query_selector('tbody tr:first-child [title*="via"]')
            )

            if not btn_2avia:
                # Fallback: procura qualquer botao/link de download na tabela
                btn_2avia = page.query_selector('tbody tr:first-child a[download], tbody tr:first-child button[download]')

            if not btn_2avia:
                resultado["msg"] = "Icone 2a Via nao encontrado na tabela de resultados"
                return resultado

            if debug:
                print(f"  Clicando 2a Via...")

            btn_2avia.click()
            time.sleep(5)

        else:
            # Sem modal: certidao foi emitida diretamente ou outro fluxo
            if debug:
                print("  Sem modal — verificando download/nova pagina direta")
            time.sleep(5)

        # ---- 7. Captura o resultado (nova pagina ou download) ----
        alvo = None

        # Verifica nova pagina
        for np in novas_paginas:
            try:
                if not np.is_closed():
                    np.wait_for_load_state("networkidle", timeout=10000)
                    alvo = np
                    break
            except Exception:
                pass

        # Se download foi detectado
        if downloads and not alvo:
            dl = downloads[0]
            dl.save_as(str(output_path))
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (download)"
            return resultado

        # Se nova pagina
        if alvo:
            if debug:
                print(f"  Nova pagina URL: {alvo.url}")
            try:
                alvo.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            alvo.pdf(path=str(output_path), format="A4", print_background=True)
            try:
                alvo.close()
            except Exception:
                pass
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (nova pagina)"
            return resultado

        # Verifica se a propria pagina principal exibe a certidao
        cur_url = page.url.lower()
        if "certidao" in cur_url and "resultado" not in cur_url:
            if debug:
                print(f"  Pagina principal tem certidao: {page.url}")
            page.pdf(path=str(output_path), format="A4", print_background=True)
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (pagina principal)"
            return resultado

        resultado["msg"] = "Certidao nao capturada (sem download nem nova pagina)"

    except Exception as e:
        resultado["msg"] = str(e)
    finally:
        try:
            context.remove_listener("page", _on_new_page)
        except Exception:
            pass
        try:
            page.remove_listener("download", _on_download)
        except Exception:
            pass

    return resultado


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--empresa",  default="EDN")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--output",   default="")
    ap.add_argument("--apenas",   default="")
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

    lista = empresa["matrizes"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]

    print(f"\n{'='*60}")
    print(f"  CND Federal (RFB) — {empresa['nome']}")
    print(f"  Total: {len(lista)} matrizes  |  Saida: {pasta}")
    print(f"{'='*60}\n")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright nao instalado.")
        sys.exit(1)

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=80)
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
