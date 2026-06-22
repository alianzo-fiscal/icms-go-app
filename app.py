# coding: utf-8
"""
app.py — Plataforma Web de Análise Fiscal ICMS/GO
Streamlit Community Cloud
"""

import os
import sys
import json
import shutil
import subprocess
import tempfile
import traceback
from pathlib import Path

import streamlit as st

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="ICMS/GO — Análise Fiscal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Autenticação simples ──────────────────────────────────────────────────────
def _senha_correta(senha_digitada: str) -> bool:
    try:
        senha_correta = st.secrets["senha"]
    except Exception:
        senha_correta = "icms2026"
    return senha_digitada == senha_correta


def tela_login():
    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown("## 🔐 Acesso Restrito")
        st.markdown("Informe a senha para acessar a plataforma.")
        senha = st.text_input("Senha", type="password", key="login_senha")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if _senha_correta(senha):
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")


if not st.session_state.get("autenticado"):
    tela_login()
    st.stop()

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://img.shields.io/badge/Alianzo-Consultoria%20Fiscal-1F3864?style=for-the-badge",
        use_container_width=True,
    )
    st.markdown("## 📊 ICMS/GO — Análise Fiscal")
    st.markdown(
        """
**Plataforma de análise tributária** para empresas varejistas no Estado de Goiás.

---

### 📋 Como usar

1. **Selecione** a aba correspondente à análise desejada
2. **Faça upload** dos arquivos XLS/CSV de movimentação
3. **Clique** em *Processar*
4. **Baixe** os relatórios gerados

---

### 📁 Formatos aceitos
- `.xls` — Excel legado (ERP)
- `.xlsx` — Excel moderno
- `.csv` — separado por `;` ou `,`

---

### ⚠️ Atenção
Os arquivos são processados em memória e **não são armazenados** no servidor.

---
"""
    )
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# ── Helpers ───────────────────────────────────────────────────────────────────
SCRIPTS_DIR = Path(__file__).parent

def _salvar_arquivo(uploaded_file, destino: Path) -> Path:
    """Grava bytes do UploadedFile em disco e retorna o Path."""
    p = destino / uploaded_file.name
    p.write_bytes(uploaded_file.getvalue())
    return p


def _formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ── Aba 1 — Análise de Entradas ───────────────────────────────────────────────
def processar_entradas(uploaded_files):
    """Chama analisar_entradas.py diretamente (importação de módulo)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Salvar arquivos
        caminhos = [str(_salvar_arquivo(f, tmpdir)) for f in uploaded_files]

        # Importar módulo do mesmo diretório
        if str(SCRIPTS_DIR) not in sys.path:
            sys.path.insert(0, str(SCRIPTS_DIR))

        import importlib
        import importlib.util

        spec = importlib.util.spec_from_file_location(
            "analisar_entradas", SCRIPTS_DIR / "analisar_entradas.py"
        )
        ae = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(ae)

        df = ae.carregar_dados(caminhos)
        periodo = ae.extrair_periodo(df, caminhos)
        inter = df[df["TIPO_OP"] == "Interestadual"]

        divs = {
            "DIV1": ae.calcular_div1(df, inter),
            "DIV2": ae.calcular_div2(inter),
            "DIV3": ae.calcular_div3(inter),
            "DIV4": ae.calcular_div4(inter),
            "DIV5": ae.calcular_div5(inter),
        }

        nome_base = f"Analise Entradas ICMS GO - {periodo}"
        excel_path = tmpdir / f"{nome_base}.xlsx"
        word_path  = tmpdir / f"{nome_base}.docx"

        ae.gerar_excel(df, divs, periodo, excel_path)
        ae.gerar_word(df, divs, periodo, word_path)

        contagens = {k: len(v) if v is not None else 0 for k, v in divs.items()}
        total_registros = len(df)
        total_inter = len(inter)

        return (
            excel_path.read_bytes(),
            word_path.read_bytes(),
            periodo,
            contagens,
            total_registros,
            total_inter,
            nome_base,
        )


# ── Aba 2 — Análise de Saídas ─────────────────────────────────────────────────
def processar_saidas(uploaded_files):
    """Chama analisar_saidas.py diretamente (importação de módulo)."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        caminhos = [str(_salvar_arquivo(f, tmpdir)) for f in uploaded_files]

        import importlib.util

        spec = importlib.util.spec_from_file_location(
            "analisar_saidas", SCRIPTS_DIR / "analisar_saidas.py"
        )
        as_ = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(as_)

        df = as_.carregar_dados(caminhos)
        periodo = as_.extrair_periodo(df, caminhos)
        intra = df[df["TIPO_OP"] == "Intraestadual"]
        inter = df[df["TIPO_OP"] == "Interestadual"]

        divs = {
            "DIV1": as_.calcular_div1(df),
            "DIV2": as_.calcular_div2(intra),
            "DIV3": as_.calcular_div3(intra),
            "DIV4": as_.calcular_div4(intra),
            "DIV5": as_.calcular_div5(df),
            "DIV6": as_.calcular_div6(inter),
            "DIV7": as_.calcular_div7(df),
        }

        grp = as_.calcular_base_consolidada(df)

        nome_base = f"Analise Saidas ICMS GO - {periodo}"
        excel_path = tmpdir / f"{nome_base}.xlsx"
        word_path  = tmpdir / f"{nome_base}.docx"

        as_.gerar_excel(df, divs, grp, periodo, excel_path)
        as_.gerar_word(df, divs, periodo, word_path)

        contagens = {k: len(v) if v is not None else 0 for k, v in divs.items()}
        total_registros = len(df)

        return (
            excel_path.read_bytes(),
            word_path.read_bytes(),
            periodo,
            contagens,
            total_registros,
            nome_base,
        )


# ── Aba 3 — Apuração Mensal ───────────────────────────────────────────────────
def processar_apuracao(ent_files, sai_files):
    """
    Chama apuracao_3abas.py via subprocess, combinando os arquivos temporários
    com combinar_xlsx.py (mesma pasta que apuracao_3abas.py).
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # Salvar arquivos de entrada e saída
        ent_paths = [str(_salvar_arquivo(f, tmpdir)) for f in ent_files]
        sai_paths = [str(_salvar_arquivo(f, tmpdir)) for f in sai_files]

        output_path = tmpdir / "Apuracao_ICMS_GO.xlsx"

        # Localizar scripts
        script_apur     = SCRIPTS_DIR / "apuracao_3abas.py"
        script_combinar = SCRIPTS_DIR / "combinar_xlsx.py"

        if not script_apur.exists():
            raise FileNotFoundError(
                f"apuracao_3abas.py não encontrado em {SCRIPTS_DIR}. "
                "Certifique-se de que todos os scripts estão na mesma pasta que app.py."
            )

        # Nomes dos arquivos temporários gerados pelo apuracao_3abas.py
        tmp_apur = tmpdir / "Apuracao_ICMS_GO_tmp_apur.xlsx"
        tmp_base = tmpdir / "Apuracao_ICMS_GO_tmp_base.xlsx"

        # Montar comando — apuracao_3abas.py vai gerar tmp_apur e tmp_base
        # A combinação será feita aqui no app.py diretamente (via importação do módulo),
        # sem depender do caminho hardcoded dentro de apuracao_3abas.py.
        cmd = (
            [sys.executable, str(script_apur)]
            + ["--entradas"] + ent_paths
            + ["--saidas"]   + sai_paths
            + ["--output",   str(output_path)]
        )

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            cwd=str(tmpdir),
        )

        stdout = result.stdout or ""
        stderr = result.stderr or ""

        # ── Caso 1: apuracao_3abas.py conseguiu combinar sozinho ──────────────
        if output_path.exists():
            excel_bytes  = output_path.read_bytes()
            nome_arquivo = "Apuracao_ICMS_GO.xlsx"

        # ── Caso 2: combinar manualmente via Python (sem subprocess) ──────────
        elif tmp_apur.exists() and tmp_base.exists() and script_combinar.exists():
            import importlib.util as _ilu
            spec = _ilu.spec_from_file_location("combinar_xlsx", script_combinar)
            cx   = _ilu.module_from_spec(spec)
            spec.loader.exec_module(cx)
            cx.combinar(str(tmp_apur), str(tmp_base), str(output_path))
            if output_path.exists():
                excel_bytes  = output_path.read_bytes()
                nome_arquivo = "Apuracao_ICMS_GO.xlsx"
                stdout += "\nAbas combinadas com sucesso."
            else:
                excel_bytes  = tmp_apur.read_bytes()
                nome_arquivo = "Apuracao_ICMS_GO_apuracao.xlsx"
                stdout += "\n⚠️  Bases consolidadas geradas em arquivo separado."

        # ── Caso 3: apenas aba de apuração disponível ─────────────────────────
        elif tmp_apur.exists():
            excel_bytes  = tmp_apur.read_bytes()
            nome_arquivo = "Apuracao_ICMS_GO_apuracao.xlsx"
            stdout += "\n⚠️  combinar_xlsx.py não encontrado. Apenas aba de apuração disponível."

        else:
            raise RuntimeError(
                f"Nenhum arquivo de saída foi gerado.\n\nSTDOUT:\n{stdout}\n\nSTDERR:\n{stderr}"
            )

        return excel_bytes, nome_arquivo, stdout, stderr


# ── Layout principal ───────────────────────────────────────────────────────────
st.title("📊 ICMS/GO — Plataforma de Análise Fiscal")
st.markdown(
    "Faça upload dos arquivos de movimentação fiscal e gere relatórios de divergências e apuração de ICMS."
)

tab_ent, tab_sai, tab_apur, tab_sped = st.tabs(
    ["📥 Análise de Entradas", "📤 Análise de Saídas", "🧮 Apuração Mensal", "📂 SPED/PVA"]
)


# ────────────────────────────────────────────────────────────────────────────
# TAB 1 — ENTRADAS
# ────────────────────────────────────────────────────────────────────────────
with tab_ent:
    st.header("Análise de Entradas — ICMS/GO")
    st.markdown(
        """
Analisa as notas fiscais de **entrada** e identifica divergências de ICMS, como:
- **DIV-1** Importada com alíquota diferente de 4% (Resolução Senado 13/2012)
- **DIV-2** CST 41 interestadual — solicitar documentação do convênio
- **DIV-3** CST 90 interestadual — verificar natureza da operação
- **DIV-4** Base de cálculo reduzida sem CST 20
- **DIV-5** Alíquota 4% + origem nacional (Res. 13/2012)

Gera relatório **Excel** (abas por divergência) + relatório **Word** (análise narrativa).
"""
    )

    st.divider()

    uploaded_ent = st.file_uploader(
        "Selecione os arquivos de **Entradas** (XLS / XLSX / CSV)",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True,
        key="uploader_entradas",
        help="Você pode selecionar múltiplos arquivos de uma vez.",
    )

    if uploaded_ent:
        st.info(f"📎 {len(uploaded_ent)} arquivo(s) selecionado(s): " + ", ".join(f.name for f in uploaded_ent))

    if uploaded_ent and st.button("▶️ Processar Entradas", type="primary", key="btn_entradas"):
        try:
            with st.spinner("Processando arquivos de entradas... Aguarde."):
                (
                    excel_bytes,
                    word_bytes,
                    periodo,
                    contagens,
                    total_registros,
                    total_inter,
                    nome_base,
                ) = processar_entradas(uploaded_ent)

            st.success(f"✅ Análise concluída — Período: **{periodo}**")

            # Métricas resumo
            total_divs = sum(contagens.values())
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total de Registros", f"{total_registros:,}")
            col2.metric("Operações Interestaduais", f"{total_inter:,}")
            col3.metric("Total de Divergências", f"{total_divs:,}")
            col4.metric("Período", periodo)

            # Detalhamento por divergência
            if total_divs > 0:
                st.markdown("#### Divergências encontradas")
                desc_divs = {
                    "DIV1": "Importada — alíquota ≠ 4%",
                    "DIV2": "CST 41 interestadual",
                    "DIV3": "CST 90 interestadual",
                    "DIV4": "Base reduzida sem CST 20",
                    "DIV5": "Alíq 4% + origem nacional",
                }
                cols = st.columns(len(contagens))
                for i, (k, v) in enumerate(contagens.items()):
                    cols[i].metric(
                        label=f"{k} — {desc_divs.get(k, k)}",
                        value=v,
                        delta="registros",
                        delta_color="off",
                    )
            else:
                st.info("Nenhuma divergência encontrada nos arquivos.")

            st.divider()
            st.markdown("#### 📥 Baixar Relatórios")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="⬇️  Baixar Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{nome_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                st.download_button(
                    label="⬇️  Baixar Word (.docx)",
                    data=word_bytes,
                    file_name=f"{nome_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        except Exception as exc:
            st.error(f"❌ Erro durante o processamento:\n\n```\n{traceback.format_exc()}\n```")


# ────────────────────────────────────────────────────────────────────────────
# TAB 2 — SAÍDAS
# ────────────────────────────────────────────────────────────────────────────
with tab_sai:
    st.header("Análise de Saídas — ICMS/GO")
    st.markdown(
        """
Analisa as notas fiscais de **saída** e identifica divergências de ICMS, como:
- **DIV-1** Alíquota 21% — atípica (verificar NCM cosméticos)
- **DIV-2** Alíquota 12% intraestadual com CST 00
- **DIV-3** Base reduzida com CST 00 (deveria ser CST 20)
- **DIV-4** CST 20 — verificar percentual de redução e estorno de crédito
- **DIV-5** CST 90 — verificar tributação pendente
- **DIV-6** CFOP 6xxx + alíquota 4% + origem nacional
- **DIV-7** CST 40/41 — confirmar convênio CONFAZ

Gera relatório **Excel** (Resumo + Base Consolidada + abas por divergência) + relatório **Word**.
"""
    )

    st.divider()

    uploaded_sai = st.file_uploader(
        "Selecione os arquivos de **Saídas** (XLS / XLSX / CSV)",
        type=["xls", "xlsx", "csv"],
        accept_multiple_files=True,
        key="uploader_saidas",
        help="Você pode selecionar múltiplos arquivos de uma vez.",
    )

    if uploaded_sai:
        st.info(f"📎 {len(uploaded_sai)} arquivo(s) selecionado(s): " + ", ".join(f.name for f in uploaded_sai))

    if uploaded_sai and st.button("▶️ Processar Saídas", type="primary", key="btn_saidas"):
        try:
            with st.spinner("Processando arquivos de saídas... Aguarde."):
                (
                    excel_bytes,
                    word_bytes,
                    periodo,
                    contagens,
                    total_registros,
                    nome_base,
                ) = processar_saidas(uploaded_sai)

            st.success(f"✅ Análise concluída — Período: **{periodo}**")

            total_divs = sum(contagens.values())
            col1, col2, col3 = st.columns(3)
            col1.metric("Total de Registros", f"{total_registros:,}")
            col2.metric("Total de Divergências", f"{total_divs:,}")
            col3.metric("Período", periodo)

            if total_divs > 0:
                st.markdown("#### Divergências encontradas")
                desc_divs = {
                    "DIV1": "Alíq 21% — atípica",
                    "DIV2": "Alíq 12% intra + CST 00",
                    "DIV3": "Base reduzida sem CST 20",
                    "DIV4": "CST 20 — verificar redução",
                    "DIV5": "CST 90 — tributação pendente",
                    "DIV6": "CFOP 6xxx + 4% + nacional",
                    "DIV7": "CST 40/41 — verificar convênio",
                }
                cols = st.columns(min(len(contagens), 4))
                for i, (k, v) in enumerate(contagens.items()):
                    cols[i % 4].metric(
                        label=f"{k} — {desc_divs.get(k, k)}",
                        value=v,
                        delta="registros",
                        delta_color="off",
                    )
            else:
                st.info("Nenhuma divergência encontrada nos arquivos.")

            st.divider()
            st.markdown("#### 📥 Baixar Relatórios")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="⬇️  Baixar Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"{nome_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                st.download_button(
                    label="⬇️  Baixar Word (.docx)",
                    data=word_bytes,
                    file_name=f"{nome_base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        except Exception as exc:
            st.error(f"❌ Erro durante o processamento:\n\n```\n{traceback.format_exc()}\n```")


# ────────────────────────────────────────────────────────────────────────────
# TAB 3 — APURAÇÃO MENSAL
# ────────────────────────────────────────────────────────────────────────────
with tab_apur:
    st.header("Apuração Mensal de ICMS — GO")
    st.markdown(
        """
Gera a planilha de **apuração de ICMS** com 3 abas:
- **APURAÇÃO ICMS** — Débito, Crédito, DIFAL Entrada/Saída, PROTEGE/GO e Saldo a Recolher por filial
- **BASE ENTRADAS** — Consolidado de entradas por filial / produto / CFOP / CST / alíquota
- **BASE SAÍDAS** — Consolidado de saídas por filial / produto / CFOP / CST / alíquota

> ℹ️ **Dica:** Envie os arquivos de entradas (XLS) e saídas (CSV/XLS) referentes ao **mesmo período**.
"""
    )

    st.divider()

    col_ent, col_sai = st.columns(2)

    with col_ent:
        st.markdown("##### 📥 Arquivo(s) de Entradas")
        uploaded_ent_apur = st.file_uploader(
            "Entradas (XLS / XLSX / CSV)",
            type=["xls", "xlsx", "csv"],
            accept_multiple_files=True,
            key="uploader_apur_ent",
        )
        if uploaded_ent_apur:
            st.caption(f"{len(uploaded_ent_apur)} arquivo(s): " + ", ".join(f.name for f in uploaded_ent_apur))

    with col_sai:
        st.markdown("##### 📤 Arquivo(s) de Saídas")
        uploaded_sai_apur = st.file_uploader(
            "Saídas (CSV / XLS / XLSX)",
            type=["xls", "xlsx", "csv"],
            accept_multiple_files=True,
            key="uploader_apur_sai",
        )
        if uploaded_sai_apur:
            st.caption(f"{len(uploaded_sai_apur)} arquivo(s): " + ", ".join(f.name for f in uploaded_sai_apur))

    pode_processar = bool(uploaded_ent_apur or uploaded_sai_apur)

    if not pode_processar:
        st.info("ℹ️  Faça upload de pelo menos um arquivo de entradas **ou** saídas para iniciar.")

    if pode_processar and st.button("▶️ Gerar Apuração", type="primary", key="btn_apuracao"):
        try:
            ent_files = uploaded_ent_apur or []
            sai_files = uploaded_sai_apur or []

            with st.spinner("Calculando apuração de ICMS... Isso pode levar alguns minutos para arquivos grandes."):
                excel_bytes, nome_arquivo, stdout, stderr = processar_apuracao(
                    ent_files, sai_files
                )

            st.success("✅ Apuração concluída com sucesso!")

            # Log do processamento (expansível)
            if stdout.strip():
                with st.expander("📋 Log do processamento", expanded=False):
                    st.code(stdout, language="text")

            if stderr.strip():
                with st.expander("⚠️  Avisos / Erros", expanded=False):
                    st.code(stderr, language="text")

            st.divider()
            st.markdown("#### 📥 Baixar Planilha de Apuração")
            st.download_button(
                label="⬇️  Baixar Excel de Apuração (.xlsx)",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except subprocess.TimeoutExpired:
            st.error("❌ O processamento excedeu o tempo limite (5 min). Tente com arquivos menores.")
        except Exception as exc:
            st.error(f"❌ Erro durante o processamento:\n\n```\n{traceback.format_exc()}\n```")

# ────────────────────────────────────────────────────────────────────────────
# TAB 4 — SPED / PVA
# ────────────────────────────
with tab_sped:
    st.subheader("📂 SPED / PVA — Validação e Transmissão em Lote")

    # ── lê config via _carregar_config() do fase1_lote (tem defaults corretos) ─
    import json as _json
    import sys as _sys
    _pva_dir = str(Path(__file__).parent / "pva_monitor")
    if _pva_dir not in _sys.path:
        _sys.path.insert(0, _pva_dir)
    try:
        from fase1_lote import _carregar_config as _cc
        _cfg           = _cc()
        _pasta_monitor = Path(_cfg["pasta_monitorada"])
        _pasta_valid   = Path(_cfg["pasta_validados"])
        _log_json_path = Path(_cfg["log_validacao"])
    except Exception as _e:
        st.error(f"Erro ao carregar configuração do PVA: {_e}")
        st.stop()

    st.caption(f"📁 Pasta monitorada: `{_pasta_monitor}`")
    st.caption(f"✅ Pasta validados:   `{_pasta_valid}`")

    st.markdown("---")

    # ── upload de TXT ─────────────────────────────────────────────────────────
    st.markdown("### 1. Upload dos arquivos TXT do SPED EFD")
    uploaded_txts = st.file_uploader(
        "Selecione um ou mais arquivos TXT gerados pelo ERP",
        type=["txt"],
        accept_multiple_files=True,
        key="uploader_sped",
    )

    if uploaded_txts:
        st.info(f"📎 {len(uploaded_txts)} arquivo(s) selecionado(s): " + ", ".join(f.name for f in uploaded_txts))

    if uploaded_txts and st.button("📥 Copiar para pasta monitorada", key="btn_copiar_sped"):
        _pasta_monitor.mkdir(parents=True, exist_ok=True)
        copiados = []
        for _f in uploaded_txts:
            _dest = _pasta_monitor / _f.name
            _dest.write_bytes(_f.getvalue())
            copiados.append(_f.name)
        st.success("Arquivo(s) copiado(s) para a pasta monitorada:\n" + "\n".join(f"- {n}" for n in copiados))

    st.markdown("---")

    # ── Fase 1: importar + validar ────────────────────────────────────────────
    st.markdown("### 2. Fase 1 — Importar e Validar no PVA")
    st.caption("O PVA Sped Fiscal será aberto automaticamente. Não interaja com o computador durante o processo.")

    if st.button("▶️ Executar Fase 1 (Importar + Validar)", type="primary", key="btn_fase1"):
        _script = Path(__file__).parent / "pva_monitor" / "fase1_lote.py"
        with st.spinner("Processando... aguarde (pode levar vários minutos por arquivo)"):
            try:
                _result = subprocess.run(
                    [sys.executable, str(_script)],
                    capture_output=True, text=True, encoding="utf-8",
                    cwd=str(_script.parent),
                    timeout=600,
                )
                _output = (_result.stdout or "") + (_result.stderr or "")
                if _result.returncode == 0:
                    st.success("✅ Fase 1 concluída!")
                else:
                    st.warning("⚠️ Fase 1 terminou com erros. Veja o log abaixo.")
                st.code(_output, language="text")
            except subprocess.TimeoutExpired:
                st.error("❌ Timeout (10 min). Verifique se o PVA está respondendo.")
            except Exception as _exc:
                st.error(f"❌ Erro: {_exc}")

    # ── resultado_validacao.json ───────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 3. Resultado da Validação")

    if _log_json_path.exists():
        try:
            _resultados = _json.loads(_log_json_path.read_text(encoding="utf-8"))
            import pandas as _pd
            _df = _pd.DataFrame(_resultados)
            if not _df.empty:
                _ok   = _df[_df.get("status", _pd.Series()) == "OK"].shape[0] if "status" in _df else 0
                _err  = _df[_df.get("status", _pd.Series()) == "ERRO"].shape[0] if "status" in _df else 0
                _f2ok = _df[_df.get("fase2_ok", _pd.Series(dtype=bool)) == True].shape[0] if "fase2_ok" in _df else 0
                col1, col2, col3 = st.columns(3)
                col1.metric("Validados OK", _ok)
                col2.metric("Com Erro",     _err)
                col3.metric("Transmitidos", _f2ok)
                st.dataframe(_df, use_container_width=True)
            else:
                st.info("Nenhum resultado registrado ainda.")
        except Exception as _exc:
            st.warning(f"Erro ao ler resultado_validacao.json: {_exc}")
    else:
        st.info("Nenhum resu