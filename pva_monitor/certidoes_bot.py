# coding: utf-8
"""
certidoes_bot.py — Abre e preenche automaticamente as certidões.

Sites automatizados (sem captcha):
  [1] CND Federal — Receita Federal (RFB + PGFN)
  [2] CRF / FGTS — Caixa Econômica Federal
  [3] CNDT — Tribunal Superior do Trabalho

Sites abertos manualmente (possuem captcha / WAF):
  [4] Certidão Estadual — SEFAZ-GO
  [5] Certidão Municipal — Prefeitura de Goiânia

Uso:
  python certidoes_bot.py --cnpj "20.758.851/0001-05"

  Ou chamado pelo app.py com os dados da empresa selecionada.
"""
import sys
import time
import argparse
from pathlib import Path

# ── Lê CNPJ via argumento ou config ──────────────────────────────────────────
def _parse_args():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cnpj",  required=True, help="CNPJ da empresa (com ou sem formatação)")
    ap.add_argument("--ie",    default="",    help="Inscrição Estadual (para SEFAZ-GO)")
    ap.add_argument("--url_municipal", default="", help="URL certidão municipal")
    return ap.parse_args()


def _so_numeros(s: str) -> str:
    return "".join(c for c in s if c.isdigit())


def _preencher_cnpj(page, cnpj_formatado: str, cnpj_numeros: str):
    """Tenta preencher o campo CNPJ na página atual."""
    seletores = [
        'input[id*="cnpj" i]',
        'input[name*="cnpj" i]',
        'input[placeholder*="CNPJ" i]',
        'input[aria-label*="CNPJ" i]',
        'input[id*="CPFCNPJ" i]',
        'input[name*="CPFCNPJ" i]',
        'input[type="text"]:visible',
    ]
    for sel in seletores:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                el.fill("")
                # Tenta formatado primeiro; se não aceitar, tenta só números
                el.type(cnpj_formatado, delay=40)
                time.sleep(0.3)
                val = el.input_value()
                if not val or val == cnpj_formatado:
                    return True
                el.fill("")
                el.type(cnpj_numeros, delay=40)
                return True
        except Exception:
            continue
    return False


def _clicar_emitir(page):
    """Tenta clicar no botão de emitir/consultar/pesquisar."""
    seletores = [
        'button[type="submit"]',
        'input[type="submit"]',
        'button:has-text("Emitir")',
        'button:has-text("Consultar")',
        'button:has-text("Pesquisar")',
        'button:has-text("Emissão")',
        'a:has-text("Emitir")',
    ]
    for sel in seletores:
        try:
            el = page.query_selector(sel)
            if el and el.is_visible():
                el.click()
                return True
        except Exception:
            continue
    return False


def main():
    args = _parse_args()
    cnpj_fmt = args.cnpj.strip()
    cnpj_num = _so_numeros(cnpj_fmt)
    ie       = args.ie.strip()
    url_mun  = args.url_municipal.strip()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Playwright não instalado.")
        print("   Execute: pip install playwright && playwright install chromium")
        sys.exit(1)

    SITES_AUTO = [
        {
            "nome"  : "CND Federal (RFB + PGFN)",
            "url"   : "https://servicos.receitafederal.gov.br/servico/certidoes/#/home/cnpj",
            "espera": 3,
        },
        {
            "nome"  : "CRF / FGTS (Caixa)",
            "url"   : "https://consulta-crf.caixa.gov.br/consultacrf/",
            "espera": 3,
        },
        {
            "nome"  : "CNDT — Débitos Trabalhistas (TST)",
            "url"   : "https://cndt-certidao.tst.jus.br/inicio.faces",
            "espera": 3,
        },
    ]

    SITES_MANUAL = [
        ("SEFAZ-GO (Certidão Estadual)", "https://www.sefaz.go.gov.br/certidao/emissao/"),
    ]
    if url_mun:
        SITES_MANUAL.append(("Certidão Municipal", url_mun))

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=50)
        context = browser.new_context()

        print(f"\nCNPJ: {cnpj_fmt}  |  IE: {ie or '—'}")
        print("=" * 55)

        # ── Sites automatizados ───────────────────────────────────
        for site in SITES_AUTO:
            print(f"\n[AUTO] {site['nome']}")
            try:
                page = context.new_page()
                page.goto(site["url"], timeout=20000, wait_until="domcontentloaded")
                time.sleep(site["espera"])
                ok = _preencher_cnpj(page, cnpj_fmt, cnpj_num)
                if ok:
                    print(f"       ✅ CNPJ preenchido")
                    _clicar_emitir(page)
                else:
                    print(f"       ⚠️  Campo não encontrado — verifique manualmente")
            except Exception as e:
                print(f"       ❌ Erro: {e}")

        # ── Sites manuais (abre, CNPJ fica na área de transferência) ──
        print("\n" + "─" * 55)
        print(f"Abrindo sites manuais. CNPJ copiado para a área de transferência.")
        try:
            context.pages[0].evaluate(f"navigator.clipboard.writeText('{cnpj_fmt}')")
        except Exception:
            pass  # clipboard pode não estar disponível em todos os contextos

        for nome, url in SITES_MANUAL:
            print(f"\n[MANUAL] {nome}")
            try:
                page = context.new_page()
                page.goto(url, timeout=20000, wait_until="domcontentloaded")
                print(f"         ✅ Aberto — cole o CNPJ manualmente (Ctrl+V)")
            except Exception as e:
                print(f"         ❌ Erro: {e}")

        print("\n" + "=" * 55)
        print("Todas as abas abertas. Feche o browser quando terminar.")
        input("Pressione Enter para fechar o browser automaticamente...")
        browser.close()


if __name__ == "__main__":
    main()
