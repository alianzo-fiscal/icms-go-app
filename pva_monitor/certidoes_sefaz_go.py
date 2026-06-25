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
            {"tag": "MATRIZ_001",  "cnpj": "20758851000105"},
            {"tag": "F003_URUACU", "cnpj": "20758851004100"},
            {"tag": "F004_FORMOSA","cnpj": "20758851004703"},
            {"tag": "F005_IPORA",  "cnpj": "20758851005610"},
            {"tag": "F006_ITUIUT", "cnpj": "20758851006120"},
            {"tag": "F007_JATAÍ",  "cnpj": "20758851006529"},
            {"tag": "F008_LUZIÂN", "cnpj": "20758851007010"},
            {"tag": "F009_MINEIROS","cnpj":"20758851007429"},
            {"tag": "F010_GOIAN",  "cnpj": "20758851007936"},
            {"tag": "F011_RIOVERD","cnpj": "20758851008410"},
            {"tag": "F012_ANÁPOL", "cnpj": "20758851008909"},
            {"tag": "F013_GOIÂNIA","cnpj": "20758851009400"},
            {"tag": "F014_GOIÂNIA","cnpj": "20758851009904"},
            {"tag": "F015_GOIÂNIA","cnpj": "20758851010508"},
            {"tag": "F016_GOIÂNIA","cnpj": "20758851011007"},
            {"tag": "F017_ITUMBIA","cnpj": "20758851011503"},
            {"tag": "F018_GOIÂNIA","cnpj": "20758851012002"},
            {"tag": "F019_CATALÃO","cnpj": "20758851012500"},
            {"tag": "F020_GOIÂNIA","cnpj": "20758851012916"},
            {"tag": "F021_GOIÂNIA","cnpj": "20758851013424"},
            {"tag": "F022_GOIÂNIA","cnpj": "20758851013912"},
            {"tag": "F023_GOIÂNIA","cnpj": "20758851014404"},
            {"tag": "F024_GOIÂNIA","cnpj": "20758851014900"},
            {"tag": "F025_GOIÂNIA","cnpj": "20758851015407"},
            {"tag": "F026_GOIÂNIA","cnpj": "20758851015903"},
            {"tag": "F027_GOIÂNIA","cnpj": "20758851016403"},
            {"tag": "F028_GOIÂNIA","cnpj": "20758851016900"},
            {"tag": "F029_GOIÂNIA","cnpj": "20758851017400"},
            {"tag": "F030_GOIÂNIA","cnpj": "20758851017906"},
            {"tag": "F031_GOIÂNIA","cnpj": "20758851018406"},
            {"tag": "F032_GOIÂNIA","cnpj": "20758851018902"},
            {"tag": "F033_GOIÂNIA","cnpj": "20758851019402"},
            {"tag": "F034_GOIÂNIA","cnpj": "20758851019909"},
            {"tag": "F035_GOIÂNIA","cnpj": "20758851020502"},
            {"tag": "F036_GOIÂNIA","cnpj": "20758851021009"},
            {"tag": "F037_GOIÂNIA","cnpj": "20758851021505"},
            {"tag": "F038_GOIÂNIA","cnpj": "20758851022006"},
            {"tag": "F039_GOIÂNIA","cnpj": "20758851022502"},
            {"tag": "F040_GOIÂNIA","cnpj": "20758851023003"},
            {"tag": "F041_GOIÂNIA","cnpj": "20758851023511"},
            {"tag": "F042_GOIÂNIA","cnpj": "20758851024003"},
            {"tag": "F043_GOIÂNIA","cnpj": "20758851024503"},
            {"tag": "F044_GOIÂNIA","cnpj": "20758851025002"},
        ],
    }
}

URL_SEFAZ = "https://www.sefaz.go.gov.br/certidao/emissao/"
DELAY_ENTRE = 4   # segundos entre cada CNPJ


def _so_numeros(s: str) -> str:
    return "".join(c for c in s if c.isdigit())


def emitir_cnpj(page, cnpj_num: str, output_path: Path) -> dict:
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

        # 3. Clica em Emitir
        page.click('input[type="submit"][value="Emitir"]', timeout=5000)
        time.sleep(2)

        # 4. Confirma o nome do contribuinte ("Sim/Não") se aparecer
        try:
            btn_sim = page.query_selector('input[value="Sim"], button:has-text("Sim")')
            if btn_sim and btn_sim.is_visible():
                btn_sim.click()
                time.sleep(3)
        except Exception:
            pass

        # 5. Aguarda a certidão final — pode abrir popup ou ficar na mesma aba
        try:
            with page.expect_popup(timeout=10000) as popup_info:
                time.sleep(1)
            nova_aba = popup_info.value
            nova_aba.wait_for_load_state("load", timeout=20000)
            time.sleep(2)
            nova_aba.pdf(path=str(output_path), print_background=True)
            nova_aba.close()
        except Exception:
            # Sem popup — certidão na mesma página
            time.sleep(3)
            page.pdf(path=str(output_path), print_background=True)

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

            res = emitir_cnpj(page, cnpj_num, pdf_path)
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
