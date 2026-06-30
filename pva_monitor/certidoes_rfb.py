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
        "nome": "Grupo Alianzo — matrizes ativas",
        "matrizes": [
            {"tag": "DN_ARMAZENAMENT", "cnpj": "28221185000183", "razao": "DN Armazenamento E Transportes Eireli"},
            {"tag": "EDN_UTILIDADES_", "cnpj": "20758851000105", "razao": "EDN Utilidades Domesticas Imp. e Exp."},
            {"tag": "R3_SUPRIMENTOS_", "cnpj": "10641901000116", "razao": "R3 Suprimentos Corporativos LTDA"},
            {"tag": "COMERCIAL_DE_BR", "cnpj": "07515610000177", "razao": "Comercial De Brinquedos Cristal Ltda"},
            {"tag": "ATACADAO_DO_LAR", "cnpj": "35917755000130", "razao": "Atacadao do Lar Com. Varejista e Atacadista"},
        ],
    },
}


def _so_numeros(s):
    return "".join(c for c in str(s) if c.isdigit())


def _formata_cnpj(cnpj_num):
    """28221185000183 -> 28.221.185/0001-83"""
    d = cnpj_num.zfill(14)
    return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"


def _pagina_em(pg, *partes):
    try:
        url = (pg.url or "").lower()
        return not pg.is_closed() and any(p in url for p in partes)
    except Exception:
        return False


def emitir_cnpj(page, context, cnpj_num, output_path, debug=False):
    """
    Fluxo RFB (requer Google Chrome via channel='chrome' para passar reCAPTCHA v3):
      1. Navega #/home/cnpj → preenche CNPJ via teclado → Emitir via JS
      2. Modal "Certidao Valida" → Consultar Certidao → /consultar
      3. /consultar → Consultar Certidao → /consultar/resultado
      4. /resultado → button[title='Segunda via'] → intercepta /seg-via/ → base64 PDF
    """
    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}
    pdf_base64 = []

    def _on_response(resp):
        if "seg-via" in resp.url and resp.status == 200:
            try:
                import base64 as _b64, json as _json
                body = resp.body()
                data = _json.loads(body)
                b64 = data.get("pdf", "")
                if b64:
                    pdf_base64.append(b64)
            except Exception:
                pass

    page.on("response", _on_response)

    # Aplica stealth para mascarar sinais de automacao (reCAPTCHA v3)
    try:
        from playwright_stealth import stealth_sync
        stealth_sync(page)
    except ImportError:
        pass  # fallback: continua sem stealth

    try:
        # 1. Navega
        page.goto(URL_RFB, timeout=60000, wait_until="load")
        time.sleep(3)
        if debug:
            print(f"  URL: {page.url}")

        # 2. Preenche CNPJ via teclado (Angular 2+ precisa de eventos reais)
        try:
            page.wait_for_selector('input[placeholder="Informe o CNPJ"]', timeout=15000)
        except Exception:
            resultado["msg"] = "Campo CNPJ nao apareceu apos 15s"
            return resultado
        time.sleep(1)

        campo_loc = page.locator('input[placeholder="Informe o CNPJ"]')
        campo_loc.click()
        page.keyboard.press("Control+a")
        page.keyboard.press("Delete")
        time.sleep(0.3)
        page.keyboard.type(cnpj_num, delay=60)
        time.sleep(0.8)

        if debug:
            print(f"  Valor preenchido: {campo_loc.input_value()}")

        # 3. Emitir via JS
        page.evaluate("""() => {
            const btn = document.querySelector('button.br-button.primary.btn-acao')
                     || document.querySelector('button[type="submit"]');
            if (btn) btn.click();
        }""")
        time.sleep(6)

        if debug:
            print(f"  Hash apos Emitir: {page.evaluate('() => window.location.hash')}")

        # Detecta erro 023
        erro_msg = page.evaluate("""() => {
            const el = document.querySelector('[class*="br-message"], [role="alert"]');
            return el ? el.innerText.trim() : null;
        }""")
        if erro_msg and any(x in erro_msg for x in ["Não foi possível", "tente novamente", "023"]):
            resultado["msg"] = f"Erro RFB 023 (reCAPTCHA) — use channel=chrome: {erro_msg[:80]}"
            return resultado

        # 4. Modal "Certidao Valida" → Consultar
        modal_ok = page.evaluate("""() => {
            return !![...document.querySelectorAll('*')].find(
                el => el.innerText && el.innerText.includes('Certidão Válida'));
        }""")

        if not modal_ok:
            resultado["msg"] = "Modal Certidao Valida nao apareceu"
            return resultado

        if debug:
            print("  Modal Certidao Valida detectado")

        # Clica Consultar Certidao (botao outline, nao o primary)
        consultado = page.evaluate("""() => {
            const btns = [...document.querySelectorAll('button')];
            const b = btns.find(b => b.innerText.includes('Consultar Certidão')
                                  && !b.classList.contains('primary'));
            if (b) { b.click(); return true; }
            const b2 = btns.find(b => b.innerText.includes('Consultar'));
            if (b2) { b2.click(); return true; }
            return false;
        }""")
        if not consultado:
            resultado["msg"] = "Botao Consultar no modal nao encontrado"
            return resultado
        time.sleep(4)

        if debug:
            print(f"  Hash /consultar: {page.evaluate('() => window.location.hash')}")

        # 5. Pagina /consultar → Consultar Certidao (primary)
        try:
            page.wait_for_url("**/consultar**", timeout=10000)
        except Exception:
            pass
        time.sleep(1)

        page.evaluate("""() => {
            const btn = document.querySelector('button.br-button.primary')
                     || [...document.querySelectorAll('button')].find(
                            b => b.innerText.includes('Consultar Certidão'));
            if (btn) btn.click();
        }""")
        time.sleep(8)

        if debug:
            print(f"  Hash /resultado: {page.evaluate('() => window.location.hash')}")

        # 6. Pagina /resultado → Segunda Via
        try:
            page.wait_for_url("**/resultado**", timeout=12000)
        except Exception:
            pass
        time.sleep(2)

        try:
            page.wait_for_selector('button[title="Segunda via"]', timeout=10000)
        except Exception:
            resultado["msg"] = "Botao Segunda via nao encontrado em /resultado"
            return resultado

        page.evaluate("""() => {
            const btn = document.querySelector('button[title="Segunda via"]');
            if (btn) btn.click();
        }""")
        time.sleep(6)

        # 7. PDF via interceptacao /seg-via/
        if pdf_base64:
            import base64
            pdf_bytes = base64.b64decode(pdf_base64[0])
            output_path.write_bytes(pdf_bytes)
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (base64 PDF via /seg-via/)"
            return resultado

        resultado["msg"] = "PDF nao capturado — resposta /seg-via/ nao interceptada"

    except Exception as e:
        resultado["msg"] = str(e)
    finally:
        try:
            page.remove_listener("response", _on_response)
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

    lista = empresa["matrizes"]
    if args.apenas:
        lista = [c for c in lista if _so_numeros(c["cnpj"]) == _so_numeros(args.apenas)]
    if args.limite > 0:
        lista = lista[:args.limite]

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
        # Usa launch_persistent_context com o perfil REAL do Chrome do usuario.
        # Isso carrega cookies, historico e sessao reCAPTCHA — essencial para passar
        # o reCAPTCHA v3 do portal RFB (que bloqueia browsers limpos com erro 023).
        import os

        # Perfil padrao do Chrome no Windows
        chrome_profile = os.path.expandvars(
            r"%LOCALAPPDATA%\Google\Chrome\User Data"
        )

        # Fecha Chrome antes de abrir com perfil persistente
        # (perfil fica travado enquanto Chrome esta aberto)
        import subprocess as _sp
        _sp.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True)
        time.sleep(2)
        print("  [Chrome] Chrome encerrado, abrindo com perfil de automacao...")

        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=chrome_profile,
                channel="chrome",
                headless=args.headless,
                slow_mo=80,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--no-restore-last-session",          # nao mostra "Restaurar paginas?"
                    "--disable-session-crashed-bubble",   # suprime popup de crash
                    "--hide-crash-restore-bubble",        # Chrome 116+
                    "--suppress-message-center-popups",
                ],
                accept_downloads=True,
            )
            # Mascara navigator.webdriver mesmo com perfil real
            context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                window.chrome = { runtime: {} };
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['pt-BR','pt','en-US','en']});
            """)
            print("  [Chrome] Usando perfil real do Chrome (com cookies/sessao)")
        except Exception as e:
            print(f"  [Chrome] Perfil ocupado ou erro ({e}) — usando contexto limpo")
            browser = p.chromium.launch(
                channel="chrome",
                headless=args.headless,
                slow_mo=80,
                args=["--disable-blink-features=AutomationControlled"],
            )
            context = browser.new_context(accept_downloads=True)
            context.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )

        pg = context.new_page()

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

        try:
            context.close()
        except Exception:
            pass

    log_path.write_text(json.dumps(resultados, ensure_ascii=False, indent=2), encoding="utf-8")
    ok  = sum(1 for r in resultados if r["status"] == "ok")
    err = sum(1 for r in resultados if r["status"] == "erro")
    print(f"\n{'='*60}")
    print(f"  Concluido: {ok} OK  |  {err} erros  |  log: {log_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
