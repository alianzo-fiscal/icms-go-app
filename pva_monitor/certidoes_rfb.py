# coding: utf-8
"""
certidoes_rfb.py — Emissao CND Federal (Receita Federal) via undetected_chromedriver.

Usa undetected_chromedriver para bypassar reCAPTCHA v3 do portal RFB.
Perfil real do Chrome do usuario e carregado automaticamente.

Uso:
  python certidoes_rfb.py --empresa EDN
  python certidoes_rfb.py --empresa GRUPO
  python certidoes_rfb.py --empresa EDN --debug
"""
import sys
import time
import argparse
import json
import os
import subprocess
from pathlib import Path
from datetime import date

URL_RFB = "https://servicos.receitafederal.gov.br/servico/certidoes/#/home/cnpj"
DELAY_ENTRE = 8

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
    d = cnpj_num.zfill(14)
    return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"


# Interceptor JS: captura resposta /seg-via/ (base64 PDF) via fetch override
FETCH_INTERCEPTOR = """
window.__rfb_pdf_b64 = null;
const __orig_fetch = window.fetch;
window.fetch = async function(...args) {
    const resp = await __orig_fetch(...args);
    try {
        const url = typeof args[0] === 'string' ? args[0] : (args[0].url || '');
        if (url.includes('seg-via')) {
            resp.clone().json().then(function(data) {
                if (data && data.pdf) { window.__rfb_pdf_b64 = data.pdf; }
            }).catch(function(){});
        }
    } catch(e) {}
    return resp;
};
"""


def emitir_cnpj(driver, cnpj_num, output_path, debug=False):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys

    resultado = {"cnpj": cnpj_num, "status": "erro", "arquivo": "", "msg": ""}

    try:
        # Limpa PDF capturado anterior
        driver.execute_script("window.__rfb_pdf_b64 = null;")

        # 1. Navega
        driver.get(URL_RFB)
        time.sleep(4)

        if debug:
            print(f"  URL: {driver.current_url}")

        # 2. Aguarda campo CNPJ
        wait = WebDriverWait(driver, 20)
        try:
            campo = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'input[placeholder="Informe o CNPJ"]')
            ))
        except Exception:
            resultado["msg"] = "Campo CNPJ nao apareceu em 20s"
            return resultado

        time.sleep(1)
        campo.click()
        time.sleep(0.3)
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        time.sleep(0.2)

        # Digita CNPJ caractere a caractere (Angular 2+)
        for char in cnpj_num:
            campo.send_keys(char)
            time.sleep(0.07)
        time.sleep(1)

        if debug:
            print(f"  Valor preenchido: {campo.get_attribute('value')}")

        # 3. Clique real no botao Emitir
        try:
            btn = driver.find_element(
                By.CSS_SELECTOR,
                "button.br-button.primary.btn-acao, button.br-button.primary"
            )
            ActionChains(driver).move_to_element(btn).pause(0.5).click().perform()
        except Exception:
            driver.execute_script("""
                const b = document.querySelector('button.br-button.primary.btn-acao')
                       || document.querySelector('button[type="submit"]');
                if (b) b.click();
            """)
        time.sleep(7)

        if debug:
            print(f"  Hash apos Emitir: {driver.execute_script('return window.location.hash')}")

        # Detecta erro 023
        try:
            msg_el = driver.find_element(By.CSS_SELECTOR, '[class*="br-message"], [role="alert"]')
            msg_txt = msg_el.text.strip()
            if any(x in msg_txt for x in ["Não foi possível", "tente novamente", "023"]):
                resultado["msg"] = f"Erro RFB 023 (reCAPTCHA): {msg_txt[:120]}"
                return resultado
        except Exception:
            pass

        # 4. Modal "Certidao Valida" → Consultar
        modal_ok = driver.execute_script("""
            return !![...document.querySelectorAll('*')].find(
                el => el.innerText && el.innerText.includes('Certidão Válida'));
        """)
        if not modal_ok:
            resultado["msg"] = "Modal Certidao Valida nao apareceu"
            return resultado

        if debug:
            print("  Modal Certidao Valida detectado")

        consultado = driver.execute_script("""
            const btns = [...document.querySelectorAll('button')];
            const b = btns.find(b => b.innerText.includes('Consultar Certidão')
                                  && !b.classList.contains('primary'));
            if (b) { b.click(); return true; }
            const b2 = btns.find(b => b.innerText.includes('Consultar'));
            if (b2) { b2.click(); return true; }
            return false;
        """)
        if not consultado:
            resultado["msg"] = "Botao Consultar no modal nao encontrado"
            return resultado
        time.sleep(5)

        # 5. Pagina /consultar → Consultar Certidao
        driver.execute_script("""
            const btn = document.querySelector('button.br-button.primary')
                     || [...document.querySelectorAll('button')].find(
                            b => b.innerText.includes('Consultar Certidão'));
            if (btn) btn.click();
        """)
        time.sleep(9)

        if debug:
            print(f"  Hash /resultado: {driver.execute_script('return window.location.hash')}")

        # 6. Pagina /resultado → Segunda Via
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'button[title="Segunda via"]')
            ))
        except Exception:
            resultado["msg"] = "Botao Segunda via nao encontrado em /resultado"
            return resultado

        driver.execute_script("""
            const btn = document.querySelector('button[title="Segunda via"]');
            if (btn) btn.click();
        """)
        time.sleep(7)

        # 7. PDF via window.__rfb_pdf_b64
        b64 = driver.execute_script("return window.__rfb_pdf_b64;")
        if b64:
            import base64
            output_path.write_bytes(base64.b64decode(b64))
            resultado["status"] = "ok"
            resultado["arquivo"] = str(output_path)
            resultado["msg"] = "OK (base64 PDF via /seg-via/)"
        else:
            resultado["msg"] = "PDF nao capturado — resposta /seg-via/ nao interceptada"

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
        import undetected_chromedriver as uc
    except ImportError as _uc_err:
        print(f"ERRO import undetected_chromedriver: {_uc_err}")
        print("Instale com:  python -m pip install undetected-chromedriver")
        sys.exit(1)

    from selenium.webdriver.common.by import By

    chrome_profile = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")

    options = uc.ChromeOptions()
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    if args.headless:
        options.add_argument("--headless=new")

    # Detecta versao do Chrome para garantir ChromeDriver compativel
    import subprocess as _sp2, re as _re
    _chrome_ver = None
    for _cpath in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]:
        if os.path.exists(_cpath):
            try:
                _out = _sp2.run([_cpath, "--version"], capture_output=True, text=True).stdout
                _m = _re.search(r"(\d+)\.", _out)
                if _m:
                    _chrome_ver = int(_m.group(1))
            except Exception:
                pass
            break
    print(f"  [Chrome] Versao detectada: {_chrome_ver}")
    print("  [Chrome] Iniciando com undetected_chromedriver...")
    driver = uc.Chrome(options=options, use_subprocess=True, version_main=_chrome_ver)
    driver.implicitly_wait(5)

    # Injeta interceptor de fetch em todas as paginas (persiste via CDP)
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": FETCH_INTERCEPTOR
        })
    except Exception as e:
        print(f"  [Aviso] Nao foi possivel instalar interceptor CDP: {e}")

    print("  [Chrome] Pronto\n")

    resultados = []
    try:
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

            res = emitir_cnpj(driver, cnpj_num, pdf_path, debug=args.debug)
            resultados.append(res)
            print("OK" if res["status"] == "ok" else f"ERRO: {res['msg']}")

            if i < len(lista):
                time.sleep(DELAY_ENTRE)

    finally:
        try:
            driver.quit()
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
