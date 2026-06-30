# coding: utf-8
"""
app.py — Alianzo Fiscal 360
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
    page_title="Alianzo | Fiscal 360",
    page_icon="https://alianzo.com.br/wp-content/uploads/2023/10/simbolo-branco.svg",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS Alianzo ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ─── Alianzo Brand Colors ─── */
:root {
    --navy:   #1B3070;
    --blue:   #1B56A3;
    --bright: #2A7DC9;
    --light:  #EAF1FB;
    --white:  #FFFFFF;
    --gray:   #F5F8FF;
    --dark:   #0D1B42;
    --border: #CDD9F0;
    --gold:   #E8A020;
    --green:  #1E9E5E;
    --red:    #D94040;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0D1B42 0%, #1B3070 60%, #1B56A3 100%) !important;
    border-right: 2px solid #2A7DC9;
}
[data-testid="stSidebar"] section > div { padding-top: 0 !important; }
[data-testid="stSidebar"] * { color: rgba(255,255,255,0.93) !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.18) !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stRadio label { color: rgba(255,255,255,0.75) !important; font-size: 0.78rem !important; }
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: rgba(255,255,255,0.12) !important;
    border-color: rgba(255,255,255,0.25) !important;
    color: white !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] [data-testid="stMarkdownContainer"] { color: white !important; }
[data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] { color: rgba(255,255,255,0.9) !important; }
[data-testid="stSidebar"] .stRadio [aria-checked="true"] + div { color: white !important; font-weight: 700 !important; }
[data-testid="stSidebar"] .stButton button {
    background: rgba(255,255,255,0.12) !important;
    border: 1px solid rgba(255,255,255,0.3) !important;
    color: white !important;
}
[data-testid="stSidebar"] .stButton button:hover {
    background: rgba(255,255,255,0.22) !important;
}

/* Top bar */
[data-testid="stHeader"] { background: var(--navy) !important; }

/* Page background */
.main .block-container { background: var(--gray); padding-top: 1.5rem !important; }

/* Header card */
.alz-header {
    background: linear-gradient(135deg, #0D1B42 0%, #1B3070 55%, #1B56A3 100%);
    border-radius: 14px;
    padding: 22px 28px 18px 28px;
    margin-bottom: 22px;
    color: white;
    display: flex;
    align-items: center;
    gap: 16px;
    box-shadow: 0 4px 20px rgba(27,48,112,0.25);
}
.alz-header .alz-icon { font-size: 2.2rem; }
.alz-header h2 { color: white !important; margin: 0; font-size: 1.5rem; font-weight: 700; }
.alz-header p  { color: rgba(255,255,255,0.75); margin: 3px 0 0 0; font-size: 0.9rem; }

/* Section headers */
.alz-section {
    background: var(--light);
    border-left: 4px solid var(--navy);
    border-radius: 0 8px 8px 0;
    padding: 9px 16px;
    margin: 20px 0 14px 0;
    font-weight: 700;
    color: var(--dark) !important;
    font-size: 0.95rem;
}

/* Info cards */
.alz-info-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 18px;
    margin: 6px 0;
    border-top: 3px solid var(--blue);
    box-shadow: 0 1px 6px rgba(27,48,112,0.07);
}
.alz-info-card .label { font-size: 0.75rem; color: #888; margin: 0; text-transform: uppercase; letter-spacing: 0.04em; }
.alz-info-card .value { font-size: 1.05rem; font-weight: 700; color: var(--dark); margin: 2px 0 0 0; }

/* Link buttons */
.alz-btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 9px 20px;
    border-radius: 8px;
    font-weight: 600;
    font-size: 0.88rem;
    text-decoration: none !important;
    transition: all 0.18s ease;
    margin: 4px 6px 4px 0;
    cursor: pointer;
}
.alz-btn-primary {
    background: var(--blue);
    color: white !important;
    border: none;
}
.alz-btn-primary:hover { background: var(--navy) !important; color: white !important; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(27,48,112,0.3); }
.alz-btn-outline {
    background: white;
    color: var(--blue) !important;
    border: 2px solid var(--blue);
}
.alz-btn-outline:hover { background: var(--light) !important; transform: translateY(-1px); }
.alz-btn-green  { background: var(--green); color: white !important; border: none; }
.alz-btn-green:hover  { background: #167a4a !important; color: white !important; transform: translateY(-1px); }
.alz-btn-gold   { background: var(--gold); color: white !important; border: none; }
.alz-btn-gold:hover   { background: #c4861a !important; color: white !important; transform: translateY(-1px); }

/* Metric override */
[data-testid="metric-container"] {
    background: white;
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 10px 14px;
    box-shadow: 0 1px 4px rgba(27,48,112,0.06);
}

/* Streamlit primary button override */
.stButton > button[kind="primary"] {
    background: var(--blue) !important;
    border-color: var(--blue) !important;
    border-radius: 8px !important;
}
.stButton > button[kind="primary"]:hover {
    background: var(--navy) !important;
    border-color: var(--navy) !important;
}
.stButton > button {
    border-radius: 8px !important;
}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    background: white;
    border-radius: 10px 10px 0 0;
    padding: 4px 8px 0 8px;
    border-bottom: 2px solid var(--border);
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0 !important;
    font-weight: 600;
    color: #666;
    padding: 8px 18px;
}
.stTabs [aria-selected="true"] {
    background: var(--light) !important;
    color: var(--navy) !important;
    border-bottom: 3px solid var(--blue) !important;
}

/* Footer */
.alz-footer {
    text-align: center;
    padding: 18px;
    color: #aaa;
    font-size: 0.78rem;
    margin-top: 40px;
    border-top: 1px solid var(--border);
}

/* Expander */
.streamlit-expanderHeader {
    background: var(--light) !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    color: var(--dark) !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: white;
    border: 2px dashed var(--border);
    border-radius: 10px;
    padding: 8px;
}

/* Divider */
hr { border-color: var(--border) !important; }
</style>
""", unsafe_allow_html=True)

# ── Autenticação simples ──────────────────────────────────────────────────────
def _senha_correta(senha_digitada: str) -> bool:
    try:
        senha_correta = st.secrets["senha"]
    except Exception:
        senha_correta = "icms2026"
    return senha_digitada == senha_correta


def tela_login():
    st.markdown("""
    <div style="display:flex;justify-content:center;margin-top:60px">
    <div style="background:white;border-radius:16px;padding:40px 48px;max-width:420px;width:100%;
                box-shadow:0 8px 40px rgba(27,48,112,0.15);border-top:4px solid #1B3070">
        <div style="text-align:center;margin-bottom:28px">
            <div style="font-size:2.4rem;margin-bottom:10px">🔒</div>
            <h2 style="color:#1B3070;margin:0 0 6px 0;font-size:1.5rem">Alianzo Fiscal 360</h2>
            <p style="color:#888;margin:0;font-size:0.9rem">Plataforma Tributária Multitributo</p>
        </div>
    </div></div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        senha = st.text_input("Senha de acesso", type="password", key="login_senha",
                              placeholder="Digite sua senha")
        if st.button("Entrar na plataforma", use_container_width=True, type="primary"):
            if _senha_correta(senha):
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
        st.markdown("<p style='text-align:center;color:#aaa;font-size:0.78rem;margin-top:16px'>Alianzo Fiscal 360 © 2026</p>",
                    unsafe_allow_html=True)


if not st.session_state.get("autenticado"):
    tela_login()
    st.stop()

# ── Dados das Empresas ────────────────────────────────────────────────────────
_EMP_LABELS = {
    "EDN":          "EDN Utilidades",
    "Atacadão":     "Atacadão do Lar",
    "Cristal":      "Comercial Cristal",
    "R3":           "R3 Suprimentos",
    "DN":           "DN Armazenamento",
    "Goyaço":       "Goyaço",
}

EMPRESAS: dict = {
    "EDN": {
        "nome_completo"          : "EDN Utilidades Domésticas Imp. e Exp. Eireli",
        "cnpj_matriz"            : "20.758.851/0001-05",
        "ie_matriz"              : "10.607.511-0",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "Varejista — Lucro Real",
        "filiais": [
            {"filial":  3, "cnpj": "20.758.851/0041-00", "ie": "20.013.121-4", "municipio": "URUAÇU",               "nome": "Uruaçu"},
            {"filial":  4, "cnpj": "20.758.851/0042-83", "ie": "20.030.770-3", "municipio": "Aparecida de Goiânia", "nome": "Mansões Paraíso"},
            {"filial":  5, "cnpj": "20.758.851/0043-64", "ie": "20.036.783-8", "municipio": "ITUMBIARA",            "nome": "Itumbiara"},
            {"filial":  7, "cnpj": "20.758.851/0045-26", "ie": "20.086.156-5", "municipio": "TRINDADE",             "nome": "Trindade"},
            {"filial":  8, "cnpj": "20.758.851/0047-98", "ie": "20.099.450-6", "municipio": "GOIÂNIA",              "nome": "Jardim Europa"},
            {"filial": 10, "cnpj": "20.758.851/0046-07", "ie": "20.092.115-0", "municipio": "CALDAS NOVAS",         "nome": "Caldas Novas"},
            {"filial": 11, "cnpj": "20.758.851/0050-93", "ie": "20.116.302-0", "municipio": "LUZIANIA",             "nome": "Luziânia"},
            {"filial": 13, "cnpj": "20.758.851/0057-60", "ie": "20.161.077-9", "municipio": "ITABERAÍ",             "nome": "Itaberaí"},
            {"filial": 16, "cnpj": "20.758.851/0019-34", "ie": "10.831.477-4", "municipio": "GOIÂNIA",              "nome": "Incorporadora 62"},
            {"filial": 18, "cnpj": "20.758.851/0022-30", "ie": "10.827.190-0", "municipio": "GOIÂNIA",              "nome": "Incorporadora 66"},
            {"filial": 20, "cnpj": "20.758.851/0034-73", "ie": "10.899.551-8", "municipio": "GOIÂNIA",              "nome": "Vila São João"},
            {"filial": 24, "cnpj": "20.758.851/0014-20", "ie": "10.827.908-1", "municipio": "GOIÂNIA",              "nome": "Incorporadora 07"},
            {"filial": 25, "cnpj": "20.758.851/0037-16", "ie": "10.940.721-0", "municipio": "INHUMAS",              "nome": "Inhumas"},
            {"filial": 26, "cnpj": "20.758.851/0056-89", "ie": "20.161.306-9", "municipio": "ANÁPOLIS",             "nome": "Anápolis 26"},
            {"filial": 27, "cnpj": "20.758.851/0051-74", "ie": "20.126.648-2", "municipio": "RIO VERDE",            "nome": "Rio Verde"},
            {"filial": 30, "cnpj": "20.758.851/0040-11", "ie": "10.987.544-3", "municipio": "GOIANIRA",             "nome": "Goianira"},
            {"filial": 31, "cnpj": "20.758.851/0001-05", "ie": "10.607.511-0", "municipio": "GOIÂNIA",              "nome": "Matriz Centro"},
            {"filial": 33, "cnpj": "20.758.851/0060-65", "ie": "20.293.165-0", "municipio": "GOIÂNIA",              "nome": "Filial 33"},
            {"filial": 38, "cnpj": "20.758.851/0058-40", "ie": "20.171.456-6", "municipio": "PORANGATU",            "nome": "Porangatu"},
            {"filial": 48, "cnpj": "20.758.851/0002-96", "ie": "10.724.310-5", "municipio": "ANÁPOLIS",             "nome": "Anápolis 48"},
            {"filial": 50, "cnpj": "20.758.851/0035-54", "ie": "10.901.034-5", "municipio": "GOIÂNIA",              "nome": "Incorporadora 44"},
            {"filial": 60, "cnpj": "20.758.851/0036-35", "ie": "10.932.688-1", "municipio": "GOIÂNIA",              "nome": "Pq. Anhanguera"},
            {"filial": 71, "cnpj": "20.758.851/0009-62", "ie": "10.811.073-7", "municipio": "GOIÂNIA",              "nome": "Hotel Real"},
            {"filial": 74, "cnpj": "20.758.851/0012-68", "ie": "10.873.370-0", "municipio": "TRINDADE",             "nome": "Trindade 74"},
            {"filial": 76, "cnpj": "20.758.851/0007-09", "ie": "10.806.526-0", "municipio": "SENADOR CANEDO",       "nome": "Senador Canedo"},
            {"filial": 77, "cnpj": "20.758.851/0008-81", "ie": "10.807.111-1", "municipio": "Aparecida de Goiânia", "nome": "Aparecida de Goiânia"},
            {"filial": 79, "cnpj": "20.758.851/0010-04", "ie": "10.814.950-1", "municipio": "GOIÂNIA",              "nome": "Incorporadora 72"},
            {"filial": 80, "cnpj": "20.758.851/0011-87", "ie": "10.819.372-1", "municipio": "GOIANÉSIA",            "nome": "Goianésia"},
            {"filial": 82, "cnpj": "20.758.851/0023-10", "ie": "10.827.170-6", "municipio": "GOIÂNIA",              "nome": "Santa Genoveva"},
            {"filial": 83, "cnpj": "20.758.851/0025-82", "ie": "10.833.248-9", "municipio": "GOIÂNIA",              "nome": "Araguaia"},
            {"filial": 84, "cnpj": "20.758.851/0024-00", "ie": "10.832.158-4", "municipio": "GOIÂNIA",              "nome": "Incorporadora 15"},
            {"filial": 85, "cnpj": "20.758.851/0013-49", "ie": "10.828.382-8", "municipio": "GOIÂNIA",              "nome": "Moinho dos Ventos"},
            {"filial": 86, "cnpj": "20.758.851/0017-72", "ie": "10.832.210-6", "municipio": "GOIÂNIA",              "nome": "Incorporadora 33"},
            {"filial": 88, "cnpj": "20.758.851/0015-00", "ie": "10.832.142-8", "municipio": "GOIÂNIA",              "nome": "Incorporadora 47"},
            {"filial": 89, "cnpj": "20.758.851/0018-53", "ie": "10.828.043-8", "municipio": "GOIÂNIA",              "nome": "Incorporadora 28"},
            {"filial": 90, "cnpj": "20.758.851/0021-59", "ie": "10.832.205-0", "municipio": "GOIÂNIA",              "nome": "Incorporadora 68"},
            {"filial": 91, "cnpj": "20.758.851/0032-01", "ie": "10.854.569-5", "municipio": "GOIÂNIA",              "nome": "Incorporadora 40"},
            {"filial": 93, "cnpj": "20.758.851/0027-44", "ie": "10.857.722-8", "municipio": "GOIÂNIA",              "nome": "Incorporadora 51"},
            {"filial": 94, "cnpj": "20.758.851/0028-25", "ie": "10.858.991-9", "municipio": "GOIÂNIA",              "nome": "Incorporadora 69"},
            {"filial": 95, "cnpj": "20.758.851/0030-40", "ie": "10.859.183-2", "municipio": "GOIÂNIA",              "nome": "Incorporadora 65"},
            {"filial": 96, "cnpj": "20.758.851/0029-06", "ie": "10.855.552-6", "municipio": "GOIÂNIA",              "nome": "Incorporadora 41"},
            {"filial": 97, "cnpj": "20.758.851/0031-20", "ie": "10.854.061-8", "municipio": "GOIÂNIA",              "nome": "Incorporadora 42"},
            {"filial": 98, "cnpj": "20.758.851/0016-91", "ie": "10.827.173-0", "municipio": "GOIÂNIA",              "nome": "Incorporadora 43"},
        ],
    },
    "Atacadão": {
        "nome_completo"          : "Atacadão do Lar Com. Varejista e Atacadista Ltda",
        "cnpj_matriz"            : "35.917.755/0001-30",
        "ie_matriz"              : "",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "Atacadista/Varejista",
        "filiais": [],
    },
    "Cristal": {
        "nome_completo"          : "Comercial de Brinquedos Cristal Ltda",
        "cnpj_matriz"            : "07.515.610/0001-77",
        "ie_matriz"              : "",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "Varejista",
        "filiais": [],
    },
    "R3": {
        "nome_completo"          : "R3 Suprimentos Corporativos LTDA",
        "cnpj_matriz"            : "10.641.901/0001-16",
        "ie_matriz"              : "",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "",
        "filiais": [],
    },
    "DN": {
        "nome_completo"          : "DN Armazenamento e Transportes Eireli",
        "cnpj_matriz"            : "28.221.185/0001-83",
        "ie_matriz"              : "",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "Armazenamento e Transporte",
        "filiais": [],
    },
    "Goyaço": {
        "nome_completo"          : "Goyaço",
        "cnpj_matriz"            : "",
        "ie_matriz"              : "",
        "municipio_sede"         : "GOIÂNIA",
        "url_certidao_municipal" : "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "perfil_tributario"      : "",
        "filiais": [],
    },
}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    # Logo / Brand
    st.markdown("""
    <div style="padding: 20px 12px 16px 12px; text-align: center;">
        <div style="font-size:1.6rem; font-weight:800; color:white; letter-spacing:0.04em">
            ALIANZO
        </div>
        <div style="font-size:0.72rem; color:rgba(255,255,255,0.55); letter-spacing:0.12em; margin-top:2px;">
            FISCAL 360
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr style='margin:0 0 14px 0'>", unsafe_allow_html=True)

    st.markdown("<p style='font-size:0.72rem;color:rgba(255,255,255,0.55);margin:0 0 5px 0;letter-spacing:0.08em;text-transform:uppercase'>Empresa</p>", unsafe_allow_html=True)
    empresa_ativa = st.selectbox(
        "empresa_sel", list(_EMP_LABELS.keys()),
        format_func=lambda x: _EMP_LABELS[x],
        key="empresa_sel", label_visibility="collapsed",
    )

    st.markdown("<hr style='margin:12px 0'>", unsafe_allow_html=True)

    st.markdown("<p style='font-size:0.72rem;color:rgba(255,255,255,0.55);margin:0 0 8px 0;letter-spacing:0.08em;text-transform:uppercase'>Módulo</p>", unsafe_allow_html=True)
    pagina_ativa = st.radio(
        "nav_pagina",
        [
            "🔭 Imersão Fiscal",
            "📊 Análise Fiscal",
            "🧮 Apuração Mensal",
            "📂 SPED / PVA",
            "💰 Guias ICMS",
            "📜 Certidões",
            "🕸️ Malha Fina Estadual",
            "💼 PIS/COFINS",
            "📅 Agenda do Líder",
        ],
        key="nav_pagina", label_visibility="collapsed",
    )

    st.markdown("<hr style='margin:14px 0 10px 0'>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.72rem;color:rgba(255,255,255,0.45);margin:0 0 4px 0'>⚠️ Arquivos processados em memória.</p>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:0.72rem;color:rgba(255,255,255,0.35);margin:0'>Nenhum dado armazenado no servidor.</p>", unsafe_allow_html=True)
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    if st.button("🚪 Sair", use_container_width=True):
        st.session_state["autenticado"] = False
        st.rerun()

# ── Helpers ───────────────────────────────────────────────────────────────────
SCRIPTS_DIR = Path(__file__).parent

def _salvar_arquivo(uploaded_file, destino: Path) -> Path:
    p = destino / uploaded_file.name
    p.write_bytes(uploaded_file.getvalue())
    return p


def _formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _btn_link(label: str, url: str, estilo: str = "primary", icon: str = "↗"):
    cls = f"alz-btn alz-btn-{estilo}"
    st.markdown(
        f'<a href="{url}" target="_blank" rel="noopener noreferrer" class="{cls}">'
        f'{label} {icon}</a>',
        unsafe_allow_html=True,
    )


def _header(icon: str, titulo: str, subtitulo: str = ""):
    sub_html = f"<p>{subtitulo}</p>" if subtitulo else ""
    st.markdown(f"""
    <div class="alz-header">
        <div class="alz-icon">{icon}</div>
        <div>
            <h2>{titulo}</h2>
            {sub_html}
        </div>
    </div>""", unsafe_allow_html=True)


def _section(label: str):
    st.markdown(f'<div class="alz-section">{label}</div>', unsafe_allow_html=True)


# ── Estado atual ──────────────────────────────────────────────────────────────
_empresa   = st.session_state.get("empresa_sel", "EDN")
_pagina    = st.session_state.get("nav_pagina",  "📊 Análise Fiscal")
_dados_emp = EMPRESAS.get(_empresa, EMPRESAS["EDN"])

# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE PROCESSAMENTO (definidas antes das páginas para evitar NameError)
# ══════════════════════════════════════════════════════════════════════════════

def _processar_entradas(uploaded_files):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        caminhos = [str(_salvar_arquivo(f, tmpdir)) for f in uploaded_files]
        if str(SCRIPTS_DIR) not in sys.path:
            sys.path.insert(0, str(SCRIPTS_DIR))
        import importlib.util
        spec = importlib.util.spec_from_file_location("analisar_entradas", SCRIPTS_DIR / "analisar_entradas.py")
        ae = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(ae)
        df = ae.carregar_dados(caminhos)
        periodo = ae.extrair_periodo(df, caminhos)
        inter = df[df["TIPO_OP"] == "Interestadual"]
        divs = {"DIV1": ae.calcular_div1(df, inter), "DIV2": ae.calcular_div2(inter),
                "DIV3": ae.calcular_div3(inter), "DIV4": ae.calcular_div4(inter),
                "DIV5": ae.calcular_div5(inter)}
        nome_base  = f"Analise Entradas ICMS GO - {periodo}"
        excel_path = tmpdir / f"{nome_base}.xlsx"
        word_path  = tmpdir / f"{nome_base}.docx"
        ae.gerar_excel(df, divs, periodo, excel_path)
        ae.gerar_word(df, divs, periodo, word_path)
        contagens = {k: len(v) if v is not None else 0 for k, v in divs.items()}
        return (excel_path.read_bytes(), word_path.read_bytes(), periodo,
                contagens, len(df), len(inter), nome_base)


def _processar_saidas(uploaded_files):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        caminhos = [str(_salvar_arquivo(f, tmpdir)) for f in uploaded_files]
        import importlib.util
        spec = importlib.util.spec_from_file_location("analisar_saidas", SCRIPTS_DIR / "analisar_saidas.py")
        as_ = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(as_)
        df = as_.carregar_dados(caminhos)
        periodo = as_.extrair_periodo(df, caminhos)
        intra = df[df["TIPO_OP"] == "Intraestadual"]
        inter = df[df["TIPO_OP"] == "Interestadual"]
        divs = {"DIV1": as_.calcular_div1(df), "DIV2": as_.calcular_div2(intra),
                "DIV3": as_.calcular_div3(intra), "DIV4": as_.calcular_div4(intra),
                "DIV5": as_.calcular_div5(df), "DIV6": as_.calcular_div6(inter),
                "DIV7": as_.calcular_div7(df)}
        grp = as_.calcular_base_consolidada(df)
        nome_base  = f"Analise Saidas ICMS GO - {periodo}"
        excel_path = tmpdir / f"{nome_base}.xlsx"
        word_path  = tmpdir / f"{nome_base}.docx"
        as_.gerar_excel(df, divs, grp, periodo, excel_path)
        as_.gerar_word(df, divs, periodo, word_path)
        contagens = {k: len(v) if v is not None else 0 for k, v in divs.items()}
        return (excel_path.read_bytes(), word_path.read_bytes(), periodo,
                contagens, len(df), nome_base)


def _processar_apuracao(ent_files, sai_files):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        ent_paths = [str(_salvar_arquivo(f, tmpdir)) for f in ent_files]
        sai_paths = [str(_salvar_arquivo(f, tmpdir)) for f in sai_files]
        output_path     = tmpdir / "Apuracao_ICMS_GO.xlsx"
        script_apur     = SCRIPTS_DIR / "apuracao_3abas.py"
        script_combinar = SCRIPTS_DIR / "combinar_xlsx.py"
        if not script_apur.exists():
            raise FileNotFoundError(f"apuracao_3abas.py não encontrado em {SCRIPTS_DIR}.")
        tmp_apur = tmpdir / "Apuracao_ICMS_GO_tmp_apur.xlsx"
        tmp_base = tmpdir / "Apuracao_ICMS_GO_tmp_base.xlsx"
        cmd = ([sys.executable, str(script_apur)]
               + ["--entradas"] + ent_paths
               + ["--saidas"]   + sai_paths
               + ["--output",   str(output_path)])
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=900, cwd=str(tmpdir))
        stdout = result.stdout or ""
        stderr = result.stderr or ""
        if output_path.exists():
            return output_path.read_bytes(), "Apuracao_ICMS_GO.xlsx", stdout, stderr
        elif tmp_apur.exists() and tmp_base.exists() and script_combinar.exists():
            import importlib.util as _ilu
            spec = _ilu.spec_from_file_location("combinar_xlsx", script_combinar)
            cx = _ilu.module_from_spec(spec); spec.loader.exec_module(cx)
            cx.combinar(str(tmp_apur), str(tmp_base), str(output_path))
            if output_path.exists():
                return output_path.read_bytes(), "Apuracao_ICMS_GO.xlsx", stdout + "\nAbas combinadas.", stderr
            return tmp_apur.read_bytes(), "Apuracao_ICMS_GO_apuracao.xlsx", stdout, stderr
        elif tmp_apur.exists():
            return tmp_apur.read_bytes(), "Apuracao_ICMS_GO_apuracao.xlsx", stdout, stderr
        raise RuntimeError(f"Nenhum arquivo gerado.\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Análise Fiscal
# ══════════════════════════════════════════════════════════════════════════════
if _pagina == "📊 Análise Fiscal":
    _header("📊", f"Análise Fiscal — {_EMP_LABELS.get(_empresa, _empresa)}",
            _dados_emp["nome_completo"])

    tab_ent, tab_sai = st.tabs(["📥 Análise de Entradas", "📤 Análise de Saídas"])

    # ── Entradas ──────────────────────────────────────────────────────────────
    with tab_ent:
        _section("📥 Entradas — Identificação de Divergências ICMS")
        st.markdown("""
Analisa as notas fiscais de **entrada** e identifica divergências de ICMS:
- **DIV-1** Importada com alíquota diferente de 4% (Res. Senado 13/2012)
- **DIV-2** CST 41 interestadual — solicitar documentação do convênio
- **DIV-3** CST 90 interestadual — verificar natureza da operação
- **DIV-4** Base de cálculo reduzida sem CST 20
- **DIV-5** Alíquota 4% + origem nacional (Res. 13/2012)

> Gera relatório **Excel** (abas por divergência) + relatório **Word** (análise narrativa).
""")
        st.divider()
        uploaded_ent = st.file_uploader(
            "Selecione os arquivos de **Entradas** (XLS / XLSX / CSV)",
            type=["xls", "xlsx", "csv"], accept_multiple_files=True,
            key="uploader_entradas",
            help="Você pode selecionar múltiplos arquivos de uma vez.",
        )
        if uploaded_ent:
            st.info(f"📎 {len(uploaded_ent)} arquivo(s): " + ", ".join(f.name for f in uploaded_ent))

        if uploaded_ent and st.button("▶️ Processar Entradas", type="primary", key="btn_entradas"):
            try:
                with st.spinner("Processando arquivos de entradas..."):
                    (excel_bytes, word_bytes, periodo, contagens,
                     total_registros, total_inter, nome_base) = _processar_entradas(uploaded_ent)
                st.success(f"✅ Análise concluída — Período: **{periodo}**")
                total_divs = sum(contagens.values())
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total de Registros", f"{total_registros:,}")
                col2.metric("Interestaduais", f"{total_inter:,}")
                col3.metric("Divergências", f"{total_divs:,}")
                col4.metric("Período", periodo)
                if total_divs > 0:
                    st.markdown("#### Divergências encontradas")
                    desc = {"DIV1":"Importada — alíq ≠ 4%","DIV2":"CST 41 interestadual",
                            "DIV3":"CST 90 interestadual","DIV4":"Base reduzida s/ CST 20",
                            "DIV5":"Alíq 4% + origem nacional"}
                    cols = st.columns(len(contagens))
                    for i, (k, v) in enumerate(contagens.items()):
                        cols[i].metric(f"{k} — {desc.get(k, k)}", v, "registros", delta_color="off")
                else:
                    st.info("Nenhuma divergência encontrada.")
                st.divider()
                _section("📥 Baixar Relatórios")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇️ Baixar Excel (.xlsx)", excel_bytes,
                                       f"{nome_base}.xlsx",
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                with c2:
                    st.download_button("⬇️ Baixar Word (.docx)", word_bytes,
                                       f"{nome_base}.docx",
                                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                       use_container_width=True)
            except Exception:
                st.error(f"❌ Erro durante o processamento:\n\n```\n{traceback.format_exc()}\n```")

    # ── Saídas ────────────────────────────────────────────────────────────────
    with tab_sai:
        _section("📤 Saídas — Identificação de Divergências ICMS")
        st.markdown("""
Analisa as notas fiscais de **saída** e identifica divergências:
- **DIV-1** Alíquota 21% — atípica (verificar NCM cosméticos)
- **DIV-2** Alíquota 12% intraestadual com CST 00
- **DIV-3** Base reduzida com CST 00 (deveria ser CST 20)
- **DIV-4** CST 20 — verificar percentual de redução e estorno de crédito
- **DIV-5** CST 90 — verificar tributação pendente
- **DIV-6** CFOP 6xxx + alíquota 4% + origem nacional
- **DIV-7** CST 40/41 — confirmar convênio CONFAZ

> Gera **Excel** (Resumo + Base Consolidada + abas por divergência) + relatório **Word**.
""")
        st.divider()
        uploaded_sai = st.file_uploader(
            "Selecione os arquivos de **Saídas** (XLS / XLSX / CSV)",
            type=["xls", "xlsx", "csv"], accept_multiple_files=True,
            key="uploader_saidas",
        )
        if uploaded_sai:
            st.info(f"📎 {len(uploaded_sai)} arquivo(s): " + ", ".join(f.name for f in uploaded_sai))

        if uploaded_sai and st.button("▶️ Processar Saídas", type="primary", key="btn_saidas"):
            try:
                with st.spinner("Processando arquivos de saídas..."):
                    (excel_bytes, word_bytes, periodo, contagens,
                     total_registros, nome_base) = _processar_saidas(uploaded_sai)
                st.success(f"✅ Análise concluída — Período: **{periodo}**")
                total_divs = sum(contagens.values())
                col1, col2, col3 = st.columns(3)
                col1.metric("Total de Registros", f"{total_registros:,}")
                col2.metric("Divergências", f"{total_divs:,}")
                col3.metric("Período", periodo)
                if total_divs > 0:
                    st.markdown("#### Divergências encontradas")
                    desc = {"DIV1":"Alíq 21% — atípica","DIV2":"Alíq 12% intra + CST 00",
                            "DIV3":"Base reduzida s/ CST 20","DIV4":"CST 20 — verificar redução",
                            "DIV5":"CST 90 — tributação pendente","DIV6":"CFOP 6xxx + 4% + nacional",
                            "DIV7":"CST 40/41 — verificar convênio"}
                    cols = st.columns(min(len(contagens), 4))
                    for i, (k, v) in enumerate(contagens.items()):
                        cols[i % 4].metric(f"{k} — {desc.get(k, k)}", v, "registros", delta_color="off")
                else:
                    st.info("Nenhuma divergência encontrada.")
                st.divider()
                _section("📥 Baixar Relatórios")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇️ Baixar Excel (.xlsx)", excel_bytes,
                                       f"{nome_base}.xlsx",
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                with c2:
                    st.download_button("⬇️ Baixar Word (.docx)", word_bytes,
                                       f"{nome_base}.docx",
                                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                       use_container_width=True)
            except Exception:
                st.error(f"❌ Erro:\n\n```\n{traceback.format_exc()}\n```")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Apuração Mensal
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "🧮 Apuração Mensal":
    _header("🧮", f"Apuração Mensal de ICMS — {_EMP_LABELS.get(_empresa, _empresa)}",
            _dados_emp["nome_completo"])

    st.markdown("""
Gera a planilha de **apuração de ICMS** com 3 abas:
- **APURAÇÃO ICMS** — Débito, Crédito, DIFAL, PROTEGE/GO e Saldo a Recolher por filial
- **BASE ENTRADAS** — Consolidado por filial / produto / CFOP / CST / alíquota
- **BASE SAÍDAS** — Consolidado por filial / produto / CFOP / CST / alíquota

> ℹ️ Envie arquivos de entradas e saídas do **mesmo período**.
""")
    st.divider()

    col_ent, col_sai = st.columns(2)
    with col_ent:
        _section("📥 Arquivo(s) de Entradas")
        uploaded_ent_apur = st.file_uploader("Entradas (XLS/XLSX/CSV)",
                                              type=["xls","xlsx","csv"],
                                              accept_multiple_files=True,
                                              key="uploader_apur_ent")
        if uploaded_ent_apur:
            st.caption(f"{len(uploaded_ent_apur)} arquivo(s): " + ", ".join(f.name for f in uploaded_ent_apur))

    with col_sai:
        _section("📤 Arquivo(s) de Saídas")
        uploaded_sai_apur = st.file_uploader("Saídas (CSV/XLS/XLSX)",
                                              type=["xls","xlsx","csv"],
                                              accept_multiple_files=True,
                                              key="uploader_apur_sai")
        if uploaded_sai_apur:
            st.caption(f"{len(uploaded_sai_apur)} arquivo(s): " + ", ".join(f.name for f in uploaded_sai_apur))

    pode_processar = bool(uploaded_ent_apur or uploaded_sai_apur)
    if not pode_processar:
        st.info("ℹ️ Faça upload de pelo menos um arquivo de entradas **ou** saídas para iniciar.")

    if pode_processar and st.button("▶️ Gerar Apuração", type="primary", key="btn_apuracao"):
        try:
            with st.spinner("Calculando apuração de ICMS..."):
                excel_bytes, nome_arquivo, stdout, stderr = _processar_apuracao(
                    uploaded_ent_apur or [], uploaded_sai_apur or [])
            st.success("✅ Apuração concluída!")
            if stdout.strip():
                with st.expander("📋 Log do processamento"):
                    st.code(stdout, language="text")
            if stderr.strip():
                with st.expander("⚠️ Avisos / Erros"):
                    st.code(stderr, language="text")
            st.divider()
            _section("📥 Baixar Planilha de Apuração")
            st.download_button("⬇️ Baixar Excel de Apuração (.xlsx)", excel_bytes,
                               nome_arquivo,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        except subprocess.TimeoutExpired:
            st.error("❌ Timeout (5 min). Tente com arquivos menores.")
        except Exception:
            st.error(f"❌ Erro:\n\n```\n{traceback.format_exc()}\n```")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: SPED / PVA
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "📂 SPED / PVA":
    _header("📂", "SPED / PVA — Validação e Transmissão em Lote",
            "EFD ICMS/IPI — automação de importação, validação e transmissão")

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

    col_a, col_b = st.columns(2)
    col_a.info(f"📁 **Pasta monitorada:** `{_pasta_monitor}`")
    col_b.info(f"✅ **Pasta validados:** `{_pasta_valid}`")

    st.markdown("---")
    _section("1️⃣ Selecionar Arquivos TXT do SPED EFD")
    uploaded_txts = st.file_uploader("Arquivos TXT gerados pelo ERP",
                                      type=["txt"], accept_multiple_files=True,
                                      key="uploader_sped")
    if uploaded_txts:
        st.info(f"📎 {len(uploaded_txts)} arquivo(s): " + ", ".join(f.name for f in uploaded_txts))

    _pasta_monitor.mkdir(parents=True, exist_ok=True)
    _txts_na_pasta = list(_pasta_monitor.glob("*.txt"))
    if _txts_na_pasta:
        st.success(f"📂 Pasta monitorada: {len(_txts_na_pasta)} arquivo(s) — " +
                   ", ".join(t.name for t in _txts_na_pasta))
    else:
        st.caption(f"📂 Pasta monitorada vazia: `{_pasta_monitor}`")

    st.markdown("---")
    _section("2️⃣ Copiar Arquivos para a Pasta Monitorada")
    if uploaded_txts and st.button("📋 Copiar arquivos para a pasta monitorada", key="btn_copiar"):
        for _f in uploaded_txts:
            (_pasta_monitor / _f.name).write_bytes(_f.getvalue())
        st.success(f"✔ {len(uploaded_txts)} arquivo(s) copiado(s).")

    _txts_prontos = list(_pasta_monitor.glob("*.txt"))
    if _txts_prontos:
        st.info(f"📂 {len(_txts_prontos)} arquivo(s) pronto(s) na pasta.")
        st.warning("✋ **Próximo passo:** abra o PVA e importe os arquivos com ➕. Depois execute a automação abaixo.")
    elif not uploaded_txts:
        st.info("ℹ️ Faça upload de pelo menos um arquivo TXT para continuar.")

    with st.expander("🗑️ Limpar pasta monitorada"):
        st.caption("Remove os .txt após importados no PVA.")
        if st.button("Limpar .txt da pasta", key="btn_limpar_pasta"):
            _removidos = sum(1 for _arq in _pasta_monitor.glob("*.txt")
                             if not _arq.unlink() or True)
            st.success(f"{_removidos} arquivo(s) removido(s).")

    st.markdown("---")
    _section("3️⃣ Executar Automação no PVA")
    st.caption("Execute após importar os arquivos. A automação irá: verificar → assinar → transmitir em lote. Não interaja com o computador.")

    if st.button("▶️ Validar → Gerar → Assinar → Transmitir",
                 type="primary", key="btn_batch"):
        _script_batch = Path(__file__).parent / "pva_monitor" / "pva_batch.py"
        import os as _os_batch
        _env_batch = _os_batch.environ.copy()
        _env_batch["PYTHONUNBUFFERED"] = "1"
        _env_batch["PYTHONUTF8"] = "1"
        with st.spinner("Automação PVA em andamento... (pode levar 30-60 min)"):
            try:
                _result_batch = subprocess.run(
                    [sys.executable, str(_script_batch)],
                    capture_output=True, text=True, encoding="utf-8",
                    cwd=str(_script_batch.parent), timeout=14400, env=_env_batch)
                _output_batch = (_result_batch.stdout or "") + (_result_batch.stderr or "")
                if _result_batch.returncode == 0:
                    st.success("✅ Automação concluída!")
                else:
                    st.error("❌ Erro na automação. Veja o log abaixo.")
                st.code(_output_batch or "(sem saída)", language="text")
            except subprocess.TimeoutExpired:
                st.error("❌ Timeout (4h). PVA não respondeu.")
            except Exception as _exc_batch:
                st.error(f"❌ Erro: {_exc_batch}")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Certidões
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "📜 Certidões":
    _cnpj      = _dados_emp.get("cnpj_matriz", "")
    _ie        = _dados_emp.get("ie_matriz", "")
    _municipio = _dados_emp.get("municipio_sede", "GOIÂNIA")
    _url_mun   = _dados_emp.get("url_certidao_municipal", "")
    _cnpj_num  = _cnpj.replace(".", "").replace("/", "").replace("-", "") if _cnpj else ""

    _header("📜", f"Certidões — {_EMP_LABELS.get(_empresa, _empresa)}",
            _dados_emp["nome_completo"])

    # ── Info bar ──────────────────────────────────────────────────────────────
    ca, cb, cc = st.columns(3)
    with ca:
        st.markdown(f"""<div class="alz-info-card">
            <p class="label">CNPJ Matriz</p>
            <p class="value">{_cnpj or "—"}</p>
        </div>""", unsafe_allow_html=True)
    with cb:
        st.markdown(f"""<div class="alz-info-card">
            <p class="label">Inscrição Estadual</p>
            <p class="value">{_ie or "—"}</p>
        </div>""", unsafe_allow_html=True)
    with cc:
        st.markdown(f"""<div class="alz-info-card">
            <p class="label">Município Sede</p>
            <p class="value">{_municipio}</p>
        </div>""", unsafe_allow_html=True)

    if not _cnpj:
        st.warning("⚠️ CNPJ não cadastrado para esta empresa. Atualize o dicionário EMPRESAS no app.py.")

    # Botão copiar CNPJ (sem número)
    if _cnpj:
        import streamlit.components.v1 as _comp_clip
        import json as _jc
        _comp_clip.html(
            "<!DOCTYPE html><html><body style='margin:6px 0 0 0'>"
            "<button id='cb' style='padding:7px 18px;background:#1B56A3;color:white;"
            "border:none;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer'>"
            "📋 Copiar CNPJ (sem pontuação)</button>"
            "<script>"
            "var cnpj=" + _jc.dumps(_cnpj_num) + ";"
            "document.getElementById('cb').addEventListener('click',function(){"
            "  navigator.clipboard.writeText(cnpj).then(function(){"
            "    document.getElementById('cb').textContent='✅ Copiado: ' + cnpj;"
            "    setTimeout(function(){document.getElementById('cb').textContent='📋 Copiar CNPJ (sem pontuação)';},2500);"
            "  });"
            "});"
            "</script></body></html>",
            height=44, scrolling=False,
        )

    # Filiais
    _filiais = _dados_emp.get("filiais", [])
    if _filiais:
        with st.expander(f"📋 Ver todas as filiais ({len(_filiais)} unidades)"):
            import pandas as _pd_cert
            _df_fil = _pd_cert.DataFrame(_filiais)[["filial","nome","cnpj","ie","municipio"]]
            _df_fil.columns = ["Filial","Nome","CNPJ","IE","Município"]
            st.dataframe(_df_fil, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── CSS extra para cards de certidão ─────────────────────────────────────
    st.markdown("""
    <style>
    .cert-card {
        background: white;
        border: 1px solid #CDD9F0;
        border-radius: 12px;
        padding: 20px 20px 16px 20px;
        border-top: 4px solid #1B56A3;
        box-shadow: 0 2px 8px rgba(27,48,112,0.08);
        height: 100%;
    }
    .cert-card .cert-icon { font-size: 1.8rem; margin-bottom: 8px; }
    .cert-card .cert-title { font-weight: 700; color: #0D1B42; font-size: 1rem; margin: 0 0 4px 0; }
    .cert-card .cert-desc  { color: #666; font-size: 0.82rem; margin: 0 0 10px 0; line-height: 1.4; }
    .cert-card .cert-cnpj  {
        background: #EAF1FB; border-radius: 6px; padding: 5px 10px;
        font-family: monospace; font-size: 0.83rem; color: #1B3070;
        margin-bottom: 12px; display: inline-block;
    }
    </style>
    """, unsafe_allow_html=True)

    _section("🏛️ Certidões Federais")

    _script_fed  = Path(__file__).parent / "pva_monitor" / "certidoes_bot.py"
    _output_fed  = Path(__file__).parent / "certidoes_output" / "Federal" / str(__import__('datetime').date.today())

    # ── Botão de automação federal (3 certidões de uma vez) ───────────────────
    if _cnpj:
        _col_auto, _col_info = st.columns([2, 3])
        with _col_auto:
            if st.button(
                f"🤖 Emitir CND Federal + FGTS + CNDT automaticamente",
                type="primary", key="btn_fed_auto",
                use_container_width=True,
            ):
                try:
                    subprocess.Popen(
                        [sys.executable, str(_script_fed),
                         "--cnpj", _cnpj, "--ie", _ie or "",
                         "--url_municipal", _url_mun or ""],
                        cwd=str(_script_fed.parent),
                    )
                    st.success(
                        f"✅ Automação iniciada para **{_EMP_LABELS.get(_empresa)}** "
                        f"(CNPJ {_cnpj}). O browser abrirá automaticamente com os 3 portais."
                    )
                except Exception as _e:
                    st.error(f"❌ {_e}")
        with _col_info:
            st.info(
                "O script abre o browser, preenche o CNPJ automaticamente e emite: "
                "**CND Federal (RFB/PGFN)**, **CRF/FGTS (Caixa)** e **CNDT (TST)**. "
                "Nenhuma interação manual necessária."
            )
    else:
        st.warning("⚠️ CNPJ não cadastrado — automação indisponível.")

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # ── 3 cards com links manuais (fallback) ──────────────────────────────────
    with st.expander("🔗 Acessar portais manualmente"):
        cf1, cf2, cf3 = st.columns(3)
        with cf1:
            st.markdown(f"""<div class="cert-card">
                <div class="cert-icon">📄</div>
                <p class="cert-title">CND Federal — RFB / PGFN</p>
                <p class="cert-desc">Certidão Negativa de Débitos Federais.</p>
                <div class="cert-cnpj">CNPJ: {_cnpj or "—"}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            _btn_link("Acessar RFB", "https://servicos.receitafederal.gov.br/servico/certidoes/#/home/cnpj", "outline")
        with cf2:
            st.markdown(f"""<div class="cert-card">
                <div class="cert-icon">🏦</div>
                <p class="cert-title">CRF — FGTS (Caixa)</p>
                <p class="cert-desc">Certificado de Regularidade do FGTS.</p>
                <div class="cert-cnpj">CNPJ: {_cnpj or "—"}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            _btn_link("Acessar Caixa", "https://consulta-crf.caixa.gov.br/consultacrf/", "outline")
        with cf3:
            st.markdown(f"""<div class="cert-card">
                <div class="cert-icon">⚖️</div>
                <p class="cert-title">CNDT — Débitos Trabalhistas</p>
                <p class="cert-desc">Certidão Negativa do TST.</p>
                <div class="cert-cnpj">CNPJ: {_cnpj or "—"}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            _btn_link("Acessar TST", "https://cndt-certidao.tst.jus.br/inicio.faces", "outline")

    # ── PDFs já emitidos ──────────────────────────────────────────────────────
    _out_fed_dir = Path(__file__).parent / "certidoes_output" / "Federal"
    _pdfs_fed = sorted(_out_fed_dir.rglob("*.pdf")) if _out_fed_dir.exists() else []
    if _pdfs_fed:
        with st.expander(f"📁 PDFs federais emitidos ({len(_pdfs_fed)} arquivos)"):
            for _p in _pdfs_fed[-20:]:
                st.markdown(f"📄 `{_p.name}`")

    st.markdown("---")
    _section("🏛️ Certidão Estadual — Goiás")

    _script_est = Path(__file__).parent / "pva_monitor" / "certidoes_sefaz_go.py"

    # Mapa empresa → chave do script
    _EMP_SCRIPT_KEY = {"EDN": "EDN", "Atacadão": "ATACADAO", "Cristal": "CRISTAL",
                       "R3": "R3", "DN": "DN", "Goyaço": "GOYACO"}
    _script_key = _EMP_SCRIPT_KEY.get(_empresa, "")

    # ── Botão automação estadual ──────────────────────────────────────────────
    _col_est1, _col_est2 = st.columns([2, 3])
    with _col_est1:
        _btn_est_disabled = not (_script_key and _script_est.exists())
        if st.button(
            f"🤖 Emitir Estadual — todas as filiais ({len(_filiais) + 1} CNPJs)",
            type="primary", key="btn_est_auto",
            disabled=_btn_est_disabled,
            use_container_width=True,
        ):
            try:
                subprocess.Popen(
                    [sys.executable, str(_script_est),
                     "--empresa", _script_key, "--debug"],
                    cwd=str(_script_est.parent),
                )
                st.success(
                    f"✅ Automação SEFAZ-GO iniciada para **{_EMP_LABELS.get(_empresa)}** "
                    f"({len(_filiais) + 1} CNPJs). Os PDFs serão salvos em `certidoes_output/SEFAZ-GO/`."
                )
            except Exception as _e:
                st.error(f"❌ {_e}")
        if _btn_est_disabled and not _script_est.exists():
            st.caption("⚠️ Script `certidoes_sefaz_go.py` não encontrado.")
    with _col_est2:
        st.info(
            "O script abre o browser e emite a certidão estadual (SEFAZ-GO) para **cada CNPJ** "
            "da empresa selecionada, salvando os PDFs automaticamente em `certidoes_output/SEFAZ-GO/<data>/`."
        )

    # ── Fallback manual + SINTEGRA ────────────────────────────────────────────
    with st.expander("🔗 Acessar portais manualmente"):
        import streamlit.components.v1 as _comp_est
        import json as _jest
        _todas_ies = []
        if _ie:
            _todas_ies.append({"nome": _dados_emp["nome_completo"] + " (Matriz)", "ie": _ie, "cnpj": _cnpj})
        for _fil in _filiais:
            if _fil.get("ie"):
                _todas_ies.append({"nome": _fil["nome"], "ie": _fil["ie"], "cnpj": _fil.get("cnpj","")})

        ce1, ce2, _esp = st.columns([1, 1, 1])
        with ce1:
            st.markdown(f"""<div class="cert-card">
                <div class="cert-icon">🟩</div>
                <p class="cert-title">CND Estadual — SEFAZ-GO</p>
                <p class="cert-desc">Certidão de regularidade fiscal estadual.</p>
                <div class="cert-cnpj">IE Matriz: {_ie or "não cadastrada"}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            _btn_link("Emitir — Matriz", "https://www.sefaz.go.gov.br/certidao/emissao/", "primary")
        with ce2:
            st.markdown(f"""<div class="cert-card">
                <div class="cert-icon">🔍</div>
                <p class="cert-title">Consulta — SINTEGRA</p>
                <p class="cert-desc">Situação cadastral e IE no SINTEGRA.</p>
                <div class="cert-cnpj">CNPJ: {_cnpj or "não cadastrado"}</div>
            </div>""", unsafe_allow_html=True)
            st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
            _btn_link("Acessar SINTEGRA", "https://www.sintegra.gov.br/", "outline")

        if _todas_ies:
            _ies_txt  = "\\n".join(f"{r['nome']}: {r['ie']}" for r in _todas_ies)
            _ies_list = _jest.dumps(_todas_ies)
            _comp_est.html(
                "<!DOCTYPE html><html><body style='margin:6px 0 0 0;padding:0'>"
                "<div style='display:flex;gap:10px;flex-wrap:wrap'>"
                "<button id='btn_open_est' style='padding:9px 20px;background:#1B56A3;color:white;"
                "border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
                f"🏛️ Abrir portal para todas as {len(_todas_ies)} filiais</button>"
                "<button id='btn_copy_est' style='padding:9px 20px;background:white;color:#1B56A3;"
                "border:2px solid #1B56A3;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
                "📋 Copiar todas as IEs</button>"
                "</div><script>"
                "var ies=" + _ies_list + ";"
                "var txt=" + _jest.dumps(_ies_txt) + ";"
                "document.getElementById('btn_open_est').addEventListener('click',function(){"
                "  ies.forEach(function(r){window.open('https://www.sefaz.go.gov.br/certidao/emissao/','_blank');});"
                "  document.getElementById('btn_open_est').textContent='✅ ' + ies.length + ' abas abertas';"
                "});"
                "document.getElementById('btn_copy_est').addEventListener('click',function(){"
                "  navigator.clipboard.writeText(txt).then(function(){"
                "    document.getElementById('btn_copy_est').textContent='✅ Copiado!';"
                "    setTimeout(function(){document.getElementById('btn_copy_est').textContent='📋 Copiar todas as IEs';},2500);"
                "  });"
                "});"
                "</script></body></html>",
                height=52, scrolling=False,
            )

    # ── PDFs estaduais já emitidos ────────────────────────────────────────────
    _out_est_dir = Path(__file__).parent / "certidoes_output" / "SEFAZ-GO"
    _pdfs_est = sorted(_out_est_dir.rglob("*.pdf")) if _out_est_dir.exists() else []
    if _pdfs_est:
        with st.expander(f"📁 PDFs estaduais emitidos ({len(_pdfs_est)} arquivos)"):
            for _p in _pdfs_est[-30:]:
                st.markdown(f"📄 `{_p.name}`")

    # Botão "Emitir para TODAS as filiais"
    if _todas_ies:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        _ies_txt  = "\\n".join(f"{r['nome']}: {r['ie']}" for r in _todas_ies)
        _ies_list = _jest.dumps(_todas_ies)
        _comp_est.html(
            "<!DOCTYPE html><html><body style='margin:0;padding:0'>"
            "<div style='display:flex;gap:10px;flex-wrap:wrap'>"
            # Botão: abrir portal para cada filial
            "<button id='btn_open_est' style='padding:9px 20px;background:#1B56A3;color:white;"
            "border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
            f"🏛️ Abrir portal para todas as {len(_todas_ies)} filiais</button>"
            # Botão: copiar todas as IEs
            "<button id='btn_copy_est' style='padding:9px 20px;background:white;color:#1B56A3;"
            "border:2px solid #1B56A3;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
            "📋 Copiar todas as IEs</button>"
            "</div>"
            "<script>"
            "var ies=" + _ies_list + ";"
            "var txt=" + _jest.dumps(_ies_txt) + ";"
            "document.getElementById('btn_open_est').addEventListener('click',function(){"
            "  ies.forEach(function(r){"
            "    window.open('https://www.sefaz.go.gov.br/certidao/emissao/','_blank');"
            "  });"
            "  document.getElementById('btn_open_est').textContent='✅ ' + ies.length + ' abas abertas — preencha cada IE';"
            "});"
            "document.getElementById('btn_copy_est').addEventListener('click',function(){"
            "  navigator.clipboard.writeText(txt).then(function(){"
            "    document.getElementById('btn_copy_est').textContent='✅ Copiado!';"
            "    setTimeout(function(){document.getElementById('btn_copy_est').textContent='📋 Copiar todas as IEs';},2500);"
            "  });"
            "});"
            "</script></body></html>",
            height=52, scrolling=False,
        )

        with st.expander(f"📋 Ver IEs de todas as filiais ({len(_todas_ies)} unidades)"):
            import pandas as _pd_ies
            _df_ies = _pd_ies.DataFrame(_todas_ies)[["nome","ie","cnpj"]]
            _df_ies.columns = ["Filial","Inscrição Estadual","CNPJ"]
            st.dataframe(_df_ies, use_container_width=True, hide_index=True)

    st.markdown("---")
    _section(f"🏙️ Certidão Municipal — {_municipio}")

    # ── Card municipal + botão todas as filiais ───────────────────────────────
    import streamlit.components.v1 as _comp_mun
    import json as _jmun

    _mun_urls = {
        "GOIÂNIA":              "https://www.goiania.go.gov.br/sistemas/sccer/asp/sccer00300f0.asp",
        "Aparecida de Goiânia": "https://www.aparecida.go.gov.br/",
        "ANÁPOLIS":             "https://tributos.anapolis.go.gov.br/",
        "RIO VERDE":            "https://www.rioverde.go.gov.br/",
        "ITUMBIARA":            "https://www.itumbiara.go.gov.br/",
        "TRINDADE":             "https://www.trindade.go.gov.br/",
        "URUAÇU":               "https://www.uruacu.go.gov.br/",
        "CALDAS NOVAS":         "https://www.caldasnovas.go.gov.br/",
        "LUZIANIA":             "https://www.luziania.go.gov.br/",
        "SENADOR CANEDO":       "https://www.senadorcanedo.go.gov.br/",
        "GOIANÉSIA":            "https://www.goianesia.go.gov.br/",
        "GOIANIRA":             "https://www.goianira.go.gov.br/",
        "INHUMAS":              "https://www.inhumas.go.gov.br/",
        "PORANGATU":            "https://www.porangatu.go.gov.br/",
        "ITABERAÍ":             "https://www.itaberai.go.gov.br/",
    }
    _url_mun_efetiva = _url_mun or _mun_urls.get(_municipio, "")

    # Agrupa filiais por município para abrir portais únicos
    _municipios_filiais: dict = {}
    if _cnpj:
        _municipios_filiais[_municipio] = {
            "url": _url_mun_efetiva,
            "filiais": [{"nome": "Matriz", "cnpj": _cnpj}]
        }
    for _fil in _filiais:
        _mun_fil = _fil.get("municipio", "")
        _url_fil = _mun_urls.get(_mun_fil, "")
        if _mun_fil not in _municipios_filiais:
            _municipios_filiais[_mun_fil] = {"url": _url_fil, "filiais": []}
        _municipios_filiais[_mun_fil]["filiais"].append({"nome": _fil["nome"], "cnpj": _fil.get("cnpj","")})

    cm1, _esp2, _esp3 = st.columns([1, 1, 1])

    with cm1:
        st.markdown(f"""
        <div class="cert-card">
            <div class="cert-icon">🏛️</div>
            <p class="cert-title">Certidão Municipal — {_municipio}</p>
            <p class="cert-desc">Certidão de regularidade fiscal junto à Prefeitura de {_municipio}.</p>
            <div class="cert-cnpj">CNPJ: {_cnpj or "não cadastrado"}</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if _url_mun_efetiva:
            _btn_link("Emitir — Matriz", _url_mun_efetiva, "primary")
        else:
            st.caption(f"⚠️ URL da Prefeitura de {_municipio} não cadastrada.")

    # Botão "Abrir portais de todos os municípios"
    _municipios_com_url = {m: d for m, d in _municipios_filiais.items() if d["url"]}
    _total_mun = len(_municipios_com_url)
    if _total_mun > 1:
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        _urls_mun_list  = _jmun.dumps([d["url"] for d in _municipios_com_url.values()])
        _cnpjs_mun_txt  = "\n".join(
            f"{mun}: " + ", ".join(f["cnpj"] for f in d["filiais"] if f["cnpj"])
            for mun, d in _municipios_com_url.items()
        )
        _comp_mun.html(
            "<!DOCTYPE html><html><body style='margin:0;padding:0'>"
            "<div style='display:flex;gap:10px;flex-wrap:wrap'>"
            "<button id='btn_open_mun' style='padding:9px 20px;background:#1E9E5E;color:white;"
            "border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
            f"🏙️ Abrir portais dos {_total_mun} municípios</button>"
            "<button id='btn_copy_mun' style='padding:9px 20px;background:white;color:#1E9E5E;"
            "border:2px solid #1E9E5E;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer'>"
            "📋 Copiar CNPJs por município</button>"
            "</div>"
            "<script>"
            "var urls=" + _urls_mun_list + ";"
            "var txt=" + _jmun.dumps(_cnpjs_mun_txt) + ";"
            "document.getElementById('btn_open_mun').addEventListener('click',function(){"
            "  urls.forEach(function(u){ window.open(u,'_blank'); });"
            "  document.getElementById('btn_open_mun').textContent='✅ ' + urls.length + ' abas abertas';"
            "});"
            "document.getElementById('btn_copy_mun').addEventListener('click',function(){"
            "  navigator.clipboard.writeText(txt).then(function(){"
            "    document.getElementById('btn_copy_mun').textContent='✅ Copiado!';"
            "    setTimeout(function(){document.getElementById('btn_copy_mun').textContent='📋 Copiar CNPJs por município';},2500);"
            "  });"
            "});"
            "</script></body></html>",
            height=52, scrolling=False,
        )

        with st.expander(f"📋 Ver filiais por município ({len(_municipios_filiais)} municípios)"):
            for _mun_nome, _mun_data in sorted(_municipios_filiais.items()):
                _url_disp = _mun_data["url"] or "—"
                _fils_str = ", ".join(f['nome'] for f in _mun_data["filiais"])
                st.markdown(f"**{_mun_nome}** — {_fils_str}  \n`{_url_disp}`")
                st.markdown("")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Guias ICMS
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "💰 Guias ICMS":
    _cnpj      = _dados_emp.get("cnpj_matriz", "")
    _ie        = _dados_emp.get("ie_matriz", "")
    _municipio = _dados_emp.get("municipio_sede", "GOIÂNIA")

    _header("💰", f"Guias ICMS — {_EMP_LABELS.get(_empresa, _empresa)}",
            f"{_dados_emp['nome_completo']} | {_municipio} — GO")

    # Info bar
    ca, cb = st.columns(2)
    with ca:
        st.markdown(f"""<div class="alz-info-card">
            <p class="label">CNPJ / IE</p>
            <p class="value">{_cnpj or "—"} &nbsp;|&nbsp; IE: {_ie or "—"}</p>
        </div>""", unsafe_allow_html=True)
    with cb:
        st.markdown(f"""<div class="alz-info-card">
            <p class="label">Estado / Município</p>
            <p class="value">Goiás (GO) / {_municipio}</p>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")

    tab_dare, tab_difal, tab_outros = st.tabs([
        "📄 DARE-GO (ICMS Estadual)",
        "📄 DIFAL / GNRE",
        "📋 Outras Guias",
    ])

    # ── DARE-GO ──────────────────────────────────────────────────────────────
    with tab_dare:
        _section("📄 DARE-GO — Documento de Arrecadação de Receitas Estaduais")
        st.markdown(f"""
        <div class="alz-info-card">
            <p class="label">Empresa</p>
            <p class="value">{_EMP_LABELS.get(_empresa, _empresa)} &nbsp;|&nbsp; IE: {_ie or "não cadastrada"}</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**📄 Emitir DARE-GO**")
            st.caption("Emissão de guia de recolhimento de ICMS no estado de Goiás (DARE).")
            _btn_link("Emitir DARE no SEFAZ-GO",
                      "https://www.sefaz.go.gov.br/DARE/",
                      "primary", "↗")

        with c2:
            st.markdown("**🔍 Consultar DARE Emitida**")
            st.caption("Verificar situação e autenticidade de DARE emitida.")
            _btn_link("Consultar no SEFAZ-GO",
                      "https://www.sefaz.go.gov.br/DARE/consulta",
                      "outline", "↗")

        with c3:
            st.markdown("**📅 Calendário Fiscal — GO**")
            st.caption("Datas de vencimento das obrigações tributárias estaduais.")
            _btn_link("Ver Calendário SEFAZ-GO",
                      "https://www.sefaz.go.gov.br/calendario/",
                      "outline", "↗")

        st.markdown("---")
        _section("ℹ️ Como emitir o DARE-GO")
        st.markdown(f"""
**Passo a passo para emissão:**

1. Acesse o link **Emitir DARE no SEFAZ-GO** acima
2. Informe a **Inscrição Estadual**: `{_ie or "(consulte no SEFAZ-GO)"}` ou o CNPJ: `{_cnpj or "(não cadastrado)"}`
3. Selecione o tipo de receita:
   - **ICMS Normal** → Código 1100
   - **ICMS DIFAL** → Código 1400 ou 1401
   - **ICMS PROTEGE/GO** → Código 1950
   - **ICMS Antecipado** → Código 1150
4. Informe o período de referência e o valor
5. Gere o código de barras e realize o pagamento

> ⚠️ O PROTEGE/GO incide a **1,5%** sobre o valor do ICMS apurado para empresas varejistas em Goiás.
""")

        st.markdown("---")
        _section("💡 Códigos DARE-GO Mais Usados")
        import pandas as _pd_dare
        _df_dare = _pd_dare.DataFrame([
            {"Código": "1100", "Descrição": "ICMS Normal — Apuração Mensal",             "Vencimento": "Dia 20 do mês seguinte"},
            {"Código": "1150", "Descrição": "ICMS Antecipado (Entrada de Mercadoria)",    "Vencimento": "No ato da entrada ou D+1"},
            {"Código": "1400", "Descrição": "ICMS DIFAL — Diferencial de Alíquota",       "Vencimento": "Dia 20 do mês seguinte"},
            {"Código": "1401", "Descrição": "ICMS DIFAL — EC 87/2015 (Consumidor Final)", "Vencimento": "Dia 15 do mês seguinte"},
            {"Código": "1950", "Descrição": "PROTEGE/GO — Fundo de Proteção Social",      "Vencimento": "Junto com ICMS normal"},
            {"Código": "1200", "Descrição": "ICMS Substituição Tributária",               "Vencimento": "Dia 20 do mês seguinte"},
            {"Código": "9410", "Descrição": "ICMS Autuação / Auto de Infração",           "Vencimento": "Conforme notificação"},
        ])
        st.dataframe(_df_dare, use_container_width=True, hide_index=True)

    # ── DIFAL / GNRE ─────────────────────────────────────────────────────────
    with tab_difal:
        _section("📄 DIFAL — Diferencial de Alíquota / GNRE")
        st.markdown("""
O **DIFAL** (Diferencial de Alíquota) é recolhido quando há compra de outro estado (DIFAL entrada)
ou venda para consumidor final em outro estado (DIFAL saída — EC 87/2015).

A **GNRE** (Guia Nacional de Recolhimento de Tributos Estaduais) é usada para recolher
ICMS-ST e DIFAL em operações interestaduais.
""")
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**📄 GNRE Online**")
            st.caption("Emissão de GNRE para ICMS-ST e DIFAL em operações interestaduais.")
            _btn_link("Emitir GNRE",
                      "https://www.gnre.pe.gov.br/gnre/portal/consultaGNRE.do",
                      "primary", "↗")
        with c2:
            st.markdown("**📄 DARE DIFAL — SEFAZ-GO**")
            st.caption("DIFAL recolhido ao estado de Goiás (código 1400/1401).")
            _btn_link("Emitir DARE DIFAL",
                      "https://www.sefaz.go.gov.br/DARE/",
                      "outline", "↗")
        with c3:
            st.markdown("**🔍 Consultar Res. 13/2012**")
            st.caption("Regras para alíquota de 4% em operações com bens importados.")
            _btn_link("Ver Resolução 13/2012",
                      "http://legis.senado.leg.br/legislacao/ListaTextoSigen.action?norma=590965",
                      "outline", "↗")

        st.markdown("---")
        _section("📊 Alíquotas Interestaduais — Tabela GO")
        import pandas as _pd_aliq
        _df_aliq = _pd_aliq.DataFrame([
            {"Origem →  Destino":  "Sul/Sudeste → GO",    "Alíquota": "12%", "Obs": "Exceto ES que é 7%"},
            {"Origem →  Destino":  "Norte/NE/CO → GO",    "Alíquota": "12%", "Obs": "Todas as regiões"},
            {"Origem →  Destino":  "Mercadoria Importada", "Alíquota": "4%",  "Obs": "Res. Senado 13/2012"},
            {"Origem →  Destino":  "GO → Sul/Sudeste",    "Alíquota": "7%",  "Obs": ""},
            {"Origem →  Destino":  "GO → Norte/NE/CO/DF", "Alíquota": "12%", "Obs": ""},
            {"Origem →  Destino":  "Alíquota GO interna", "Alíquota": "12% / 17% / 19% / 21%",
             "Obs": "Varia por produto (RCTE/GO)"},
        ])
        st.dataframe(_df_aliq, use_container_width=True, hide_index=True)

    # ── Outras Guias ─────────────────────────────────────────────────────────
    with tab_outros:
        _section("📋 Outras Guias e Obrigações Acessórias")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🏛️ Federal")
            guias_fed = [
                ("DARF Online (RFB)",          "https://sicalc.rfb.gov.br/",             "primary",  "Recolhimento de tributos federais (PIS, COFINS, CSLL, IRPJ)"),
                ("FGTS Digital (eSocial)",      "https://fgtsdigital.caixa.gov.br/",       "outline", "Recolhimento mensal do FGTS"),
                ("GPS — Contribuição INSS",     "https://www.issa.gov.br/",                "outline", "Guia da Previdência Social"),
                ("e-CAC (Caixa Postal RFB)",    "https://cav.receita.fazenda.gov.br/",     "outline", "Acesso a mensagens, procurações e serviços RFB"),
            ]
            for label, url, estilo, descr in guias_fed:
                st.markdown(f"**{label}**")
                st.caption(descr)
                _btn_link(label, url, estilo, "↗")
                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        with c2:
            st.markdown("#### 🏛️ Estadual / GO")
            guias_est = [
                ("GIA-GO (Escrituração Fiscal)",    "https://www.sefaz.go.gov.br/GIA/",            "primary", "Geração e envio da GIA mensal à SEFAZ-GO"),
                ("EFD ICMS/IPI (SPED Fiscal)",      "https://www.receita.fazenda.gov.br/",          "outline", "Entrega da escrituração digital"),
                ("DeSTDA — Simples Nacional",        "https://destda.fazenda.gov.br/",              "outline", "Declaração para empresas do Simples Nacional"),
                ("SINTEGRA-GO (Consulta)",           "https://www.sintegra.gov.br/",                "outline", "Cadastro e situação de contribuintes no SINTEGRA"),
                ("Nota Fiscal Goiana (NF-e GO)",     "https://www.sefaz.go.gov.br/",                "outline", "Autorização e consulta de NF-e em Goiás"),
            ]
            for label, url, estilo, descr in guias_est:
                st.markdown(f"**{label}**")
                st.caption(descr)
                _btn_link(label, url, estilo, "↗")
                st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

        st.markdown("---")
        _section("📅 Calendário de Obrigações — Goiás")
        import pandas as _pd_cal
        _df_cal = _pd_cal.DataFrame([
            {"Obrigação": "ICMS Normal (DARE-GO)",          "Vencimento": "Dia 20 do mês seguinte",     "Periodicidade": "Mensal"},
            {"Obrigação": "ICMS PROTEGE/GO",                "Vencimento": "Junto com ICMS Normal",      "Periodicidade": "Mensal"},
            {"Obrigação": "ICMS DIFAL (entrada)",           "Vencimento": "Dia 20 do mês seguinte",     "Periodicidade": "Mensal"},
            {"Obrigação": "ICMS Antecipado (§ 2º, art. 13)","Vencimento": "No ato da entrada",          "Periodicidade": "Por operação"},
            {"Obrigação": "GIA-GO",                         "Vencimento": "Dia 15 do mês seguinte",     "Periodicidade": "Mensal"},
            {"Obrigação": "EFD ICMS/IPI (SPED Fiscal)",     "Vencimento": "Dia 15 do 2º mês seguinte",  "Periodicidade": "Mensal"},
            {"Obrigação": "FGTS Digital",                   "Vencimento": "Dia 20 do mês seguinte",     "Periodicidade": "Mensal"},
        ])
        st.dataframe(_df_cal, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Imersão Fiscal
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "🔭 Imersão Fiscal":
    # ── Perfis fiscais por empresa ────────────────────────────────────────────
    _PERFIS = {
        "DN": {
            "regime":       "Lucro Presumido",
            "setor":        "Transporte e Armazenamento",
            "ufs_principais": ["GO", "ES"],
            "faturamento":  "Operação limitada ao grupo",
            "icms_efetivo": "—",
            "beneficios":   "Sem benefício fiscal",
            "resumo": (
                "Operação concentrada no atendimento ao próprio grupo, sem benefício fiscal. "
                "Faturamento distribuído entre Goiás e Espírito Santo. "
                "Tributada no Lucro Presumido. Setor de transporte e armazenamento."
            ),
            "destaques": [
                ("🏢 Regime", "Lucro Presumido"),
                ("📍 UFs principais", "Goiás / Espírito Santo"),
                ("🎯 Benefício fiscal", "Nenhum"),
                ("🔄 Operação", "Exclusiva para o grupo"),
            ],
        },
        "EDN": {
            "regime":       "Lucro Real",
            "setor":        "Varejo — Lojas do Grupo BIG",
            "ufs_principais": ["GO"],
            "faturamento":  "~R$ 30 milhões/mês",
            "icms_efetivo": "14%",
            "beneficios":   "Isenção cesta básica / Redução papelaria 12%",
            "resumo": (
                "Lojas do Grupo BIG. Lucro Real. 80% da operação concentrada em Goiás. "
                "Faturamento médio de R$ 30 milhões/mês. Carga efetiva de ICMS de 14%. "
                "Produtos da cesta básica com isenção; papelaria com redução de base para alíquota efetiva de 12%."
            ),
            "destaques": [
                ("🏢 Regime", "Lucro Real"),
                ("📍 Concentração", "80% em Goiás"),
                ("💰 Faturamento médio", "~R$ 30 milhões/mês"),
                ("📊 ICMS efetivo", "14%"),
                ("✅ Cesta básica", "Isenção de ICMS"),
                ("✅ Papelaria", "Redução — alíq. efetiva 12%"),
            ],
        },
        "R3": {
            "regime":       "Lucro Real",
            "setor":        "Atacado / Suprimentos Corporativos",
            "ufs_principais": ["GO"],
            "faturamento":  "—",
            "icms_efetivo": "6%",
            "beneficios":   "Redução de base — vendas em atacado",
            "resumo": (
                "Lucro Real. Única empresa do grupo com foco em suprimentos corporativos. "
                "Boa parte das vendas beneficiadas por redução de base de cálculo em operações de atacado. "
                "Alíquota efetiva de ICMS de 6%."
            ),
            "destaques": [
                ("🏢 Regime", "Lucro Real"),
                ("📊 ICMS efetivo", "6%"),
                ("✅ Benefício", "Redução base — operações atacado"),
                ("🏭 Perfil", "Única empresa do grupo — suprimentos"),
            ],
        },
        "Cristal": {
            "regime":       "Lucro Real",
            "setor":        "Importação de Brinquedos",
            "ufs_principais": ["SC"],
            "faturamento":  "40% do abastecimento do Grupo BIG",
            "icms_efetivo": "2,6%",
            "beneficios":   "TTD Santa Catarina — carga efetiva 2,6%",
            "resumo": (
                "Lucro Real. Atividade principal de importação de brinquedos. "
                "Operações concentradas em Santa Catarina. "
                "Benefício do TTD (Tratamento Tributário Diferenciado) com carga efetiva de ICMS de 2,6%. "
                "Responsável por 40% do abastecimento do Grupo BIG."
            ),
            "destaques": [
                ("🏢 Regime", "Lucro Real"),
                ("📍 UF principal", "Santa Catarina"),
                ("📊 ICMS efetivo", "2,6%"),
                ("✅ Benefício", "TTD — SC"),
                ("📦 Papel no grupo", "40% do abastecimento BIG"),
            ],
        },
        "Atacadão": {
            "regime":       "Lucro Real",
            "setor":        "Atacado — Utensílios para o Lar",
            "ufs_principais": ["ES", "TO"],
            "faturamento":  "—",
            "icms_efetivo": "1,14% (ES) / crédito presumido 75% (TO)",
            "beneficios":   "COMPETE-ES + Crédito Presumido 75% Tocantins",
            "resumo": (
                "Lucro Real. Atacado concentrado no Espírito Santo e Tocantins. "
                "Opera via ES com carga efetiva de ICMS de 1,14% (Programa COMPETE-ES). "
                "TTD do Tocantins com crédito presumido de 75% do valor da guia."
            ),
            "destaques": [
                ("🏢 Regime", "Lucro Real"),
                ("📍 UFs principais", "Espírito Santo / Tocantins"),
                ("📊 ICMS efetivo ES", "1,14%"),
                ("✅ Benefício ES", "COMPETE-ES"),
                ("✅ Benefício TO", "Crédito presumido 75% da guia"),
            ],
        },
        "Goyaço": {
            "regime":       "—",
            "setor":        "—",
            "ufs_principais": ["GO"],
            "faturamento":  "—",
            "icms_efetivo": "—",
            "beneficios":   "—",
            "resumo":       "Perfil fiscal em cadastramento. Atualize as informações no dicionário EMPRESAS.",
            "destaques":    [("⚠️ Status", "Perfil fiscal pendente de cadastro")],
        },
    }

    _perfil = _PERFIS.get(_empresa, _PERFIS["Goyaço"])

    _header("🔭", f"Imersão Fiscal — {_EMP_LABELS.get(_empresa, _empresa)}",
            _dados_emp["nome_completo"])

    # ── Card de destaque da empresa selecionada ───────────────────────────────
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#0D1B42 0%,#1B3070 55%,#1B56A3 100%);
                border-radius:14px;padding:26px 30px;color:white;margin-bottom:20px;
                box-shadow:0 4px 20px rgba(27,48,112,0.3)">
        <div style="display:flex;align-items:flex-start;gap:20px;flex-wrap:wrap">
            <div style="flex:1;min-width:220px">
                <p style="color:rgba(255,255,255,0.55);font-size:0.75rem;margin:0 0 6px 0;
                          letter-spacing:0.1em;text-transform:uppercase">Perfil Tributário</p>
                <h3 style="color:white;margin:0 0 12px 0;font-size:1.2rem">
                    {_EMP_LABELS.get(_empresa, _empresa)}
                </h3>
                <p style="color:rgba(255,255,255,0.85);font-size:0.92rem;line-height:1.6;margin:0">
                    {_perfil["resumo"]}
                </p>
            </div>
            <div style="display:flex;flex-direction:column;gap:8px;min-width:200px">
                {''.join(f'''<div style="background:rgba(255,255,255,0.1);border-radius:8px;
                    padding:7px 14px;display:flex;justify-content:space-between;gap:16px">
                    <span style="color:rgba(255,255,255,0.6);font-size:0.78rem">{k}</span>
                    <span style="color:white;font-size:0.82rem;font-weight:700;text-align:right">{v}</span>
                </div>''' for k, v in _perfil["destaques"])}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Benefícios Fiscais — dicionário completo ──────────────────────────────
    _BENEFICIOS_LEGIS = {
        "isencao_cesta_basica_go": {
            "nome":       "Isenção — Cesta Básica (GO)",
            "uf":         "Goiás",
            "icms":       "0% (isenção total)",
            "base_legal": "RCTE/GO — Decreto 4.852/97, Anexo IX, item 19",
            "lei":        "Lei Estadual 11.651/91 (CTE-GO), art. 8º, inciso II",
            "url":        "https://www.sefaz.go.gov.br/legislacao/rcte/",
            "resumo": (
                "A legislação goiana concede **isenção de ICMS** nas saídas internas "
                "de produtos que compõem a cesta básica, como arroz, feijão, óleo de soja "
                "comestível, sal de cozinha, açúcar cristal, macarrão, farinha de mandioca e "
                "fubá de milho, entre outros elencados no Anexo IX do RCTE/GO. "
                "O benefício é condicionado à inexistência de similar produzido no Estado e "
                "ao não aproveitamento de crédito relativo às entradas (estorno proporcional). "
                "A manutenção do benefício exige que o preço ao consumidor não ultrapasse os "
                "limites fixados pela SEFAZ-GO."
            ),
        },
        "reducao_papelaria_go": {
            "nome":       "Redução de Base — Papelaria (GO)",
            "uf":         "Goiás",
            "icms":       "~12% efetivo (redução de base)",
            "base_legal": "RCTE/GO — Decreto 4.852/97, Anexo VIII",
            "lei":        "Lei Estadual 11.651/91 (CTE-GO), art. 11",
            "url":        "https://www.sefaz.go.gov.br/legislacao/rcte/",
            "resumo": (
                "O RCTE/GO prevê **redução da base de cálculo do ICMS** nas operações internas "
                "com produtos de papelaria (cadernos, canetas, lápis, borrachas, réguas e "
                "artigos escolares correlatos, classificados em NCMs específicos do Capítulo 48 "
                "e 96 da TIPI). A redução resulta em carga tributária efetiva de **12%** "
                "(equivalente a aplicar 70,59% sobre a alíquota interna padrão de 17%). "
                "O benefício é aplicável às saídas internas destinadas a varejistas e "
                "consumidores finais, vedado o aproveitamento integral do crédito — "
                "exige-se estorno proporcional à redução concedida."
            ),
        },
        "reducao_atacado_go": {
            "nome":       "Redução de Base — Atacado (GO)",
            "uf":         "Goiás",
            "icms":       "~6% efetivo",
            "base_legal": "RCTE/GO — Decreto 4.852/97, Anexo VIII; e Convênio ICMS 70/97",
            "lei":        "Lei Estadual 11.651/91 (CTE-GO), art. 11",
            "url":        "https://www.sefaz.go.gov.br/legislacao/rcte/",
            "resumo": (
                "Para operações de **saída em atacado** (vendas para revendedores com "
                "destino a comercialização), o RCTE/GO e Convênios CONFAZ aplicáveis "
                "preveem redução da base de cálculo do ICMS que resulta em carga efetiva "
                "de aproximadamente **6%** sobre o valor das mercadorias. "
                "O benefício aplica-se quando o destinatário é contribuinte do ICMS e "
                "a mercadoria se destina a posterior comercialização. "
                "Exige escrituração específica no SPED (CST 20) com indicação do "
                "percentual de redução e o correspondente estorno de crédito proporcional "
                "à parcela reduzida, conforme art. 59 do RCTE/GO."
            ),
        },
        "ttd_sc": {
            "nome":       "TTD — Importação (SC)",
            "uf":         "Santa Catarina",
            "icms":       "2,6% carga efetiva",
            "base_legal": "RICMS/SC — Decreto 2.870/01, Anexo 3, art. 10; TTD nº 409",
            "lei":        "Lei Estadual SC 10.297/96, art. 43",
            "url":        "https://www.sef.sc.gov.br/legislacao/ricms",
            "resumo": (
                "O **Tratamento Tributário Diferenciado (TTD)** concedido pelo Estado de "
                "Santa Catarina a importadores estabelecidos no estado permite o "
                "diferimento parcial do ICMS incidente nas operações de importação de "
                "mercadorias do exterior. Na prática, a empresa recolhe apenas **2,6%** "
                "de ICMS sobre o valor CIF da mercadoria importada (base de cálculo "
                "conforme art. 37, §1º do RICMS/SC), diferindo o restante do imposto "
                "para o momento da saída subsequente. "
                "O TTD é concedido pela SEF/SC mediante regime especial, exige "
                "contrapartida de geração de empregos e faturamento mínimo no estado, "
                "e deve ser renovado periodicamente conforme Portaria SEF/SC vigente."
            ),
        },
        "compete_es": {
            "nome":       "COMPETE-ES",
            "uf":         "Espírito Santo",
            "icms":       "1,14% carga efetiva",
            "base_legal": "Lei ES 10.550/2016; Decreto ES 3.974-R/2016",
            "lei":        "Lei Estadual ES 10.550/2016 (Programa COMPETE-ES)",
            "url":        "https://sefaz.es.gov.br/legislacao",
            "resumo": (
                "O **Programa COMPETE-ES** (Lei 10.550/2016) é o benefício fiscal do "
                "Espírito Santo que concede **crédito presumido de ICMS** a empresas "
                "atacadistas e distribuidoras com operações no estado. "
                "O programa permite que o contribuinte utilize crédito presumido de "
                "forma que a carga tributária efetiva nas saídas interestaduais seja "
                "reduzida para **1,14%** sobre o valor das mercadorias. "
                "É um instrumento de política industrial do ES para atrair centros "
                "de distribuição e atacadistas, amplamente utilizado por grandes "
                "varejistas e atacadistas nacionais que operam a partir do porto de "
                "Vitória (ES). Exige enquadramento prévio na SEFAZ-ES, cumprimento de "
                "metas de faturamento e geração de empregos no estado."
            ),
        },
        "credito_presumido_to": {
            "nome":       "Crédito Presumido 75% — Atacado (TO)",
            "uf":         "Tocantins",
            "icms":       "25% do ICMS nominal (75% de crédito presumido)",
            "base_legal": "RICMS/TO — Decreto 2.912/06; Instrução Normativa SEFAZ-TO",
            "lei":        "Lei Estadual TO 1.201/2000; Decreto 2.912/2006",
            "url":        "https://www.sefaz.to.gov.br/legislacao/",
            "resumo": (
                "O Tocantins concede **crédito presumido de 75%** do valor do ICMS "
                "apurado nas operações de saída para empresas atacadistas com inscrição "
                "estadual no TO. Na prática, o contribuinte recolhe apenas **25%** "
                "do ICMS nominalmente devido, sendo o restante (75%) reconhecido como "
                "crédito presumido sem necessidade de recolhimento. "
                "O benefício está previsto no RICMS/TO e em Instruções Normativas da "
                "SEFAZ-TO, sendo concedido mediante regime especial a distribuidoras "
                "e atacadistas que se instalam no estado com o objetivo de atrair "
                "centros logísticos para a região Norte/Centro-Oeste. "
                "Exige manutenção de estabelecimento físico, empregos formais e "
                "cumprimento de obrigações acessórias no TO."
            ),
        },
    }

    # Mapa empresa → benefícios
    _EMPRESA_BENEFICIOS = {
        "EDN":      ["isencao_cesta_basica_go", "reducao_papelaria_go"],
        "R3":       ["reducao_atacado_go"],
        "Cristal":  ["ttd_sc"],
        "Atacadão": ["compete_es", "credito_presumido_to"],
        "DN":       [],
        "Goyaço":   [],
    }

    _bens_empresa = _EMPRESA_BENEFICIOS.get(_empresa, [])

    if _bens_empresa:
        _section("🎯 Benefícios Fiscais — " + _EMP_LABELS.get(_empresa, _empresa))
        for _bid in _bens_empresa:
            _b = _BENEFICIOS_LEGIS[_bid]
            with st.expander(f"📋 {_b['nome']}  —  {_b['uf']}  |  ICMS efetivo: {_b['icms']}"):
                bc1, bc2 = st.columns([2, 1])
                with bc1:
                    st.markdown(f"**Resumo legislativo:**\n\n{_b['resumo']}")
                with bc2:
                    st.markdown(f"""
<div class="alz-info-card">
    <p class="label">Base Legal</p>
    <p class="value" style="font-size:0.82rem">{_b['base_legal']}</p>
</div>
<div class="alz-info-card" style="margin-top:8px">
    <p class="label">Lei / Decreto</p>
    <p class="value" style="font-size:0.82rem">{_b['lei']}</p>
</div>
""", unsafe_allow_html=True)
                    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
                    _btn_link("Ver legislação oficial", _b["url"], "outline", "↗")
    else:
        _section("🎯 Benefícios Fiscais")
        st.info(f"Nenhum benefício fiscal cadastrado para {_EMP_LABELS.get(_empresa, _empresa)}.")

    st.markdown("---")

    # ── Painel comparativo ────────────────────────────────────────────────────
    _section("📊 Resumo Comparativo — Todas as Empresas do Grupo")

    import pandas as _pd_im
    _df_im = _pd_im.DataFrame([
        {
            "Empresa":          _EMP_LABELS[k],
            "Regime":           _PERFIS[k]["regime"],
            "Setor":            _PERFIS[k]["setor"],
            "UFs Principais":   " / ".join(_PERFIS[k]["ufs_principais"]),
            "ICMS Efetivo":     _PERFIS[k]["icms_efetivo"],
            "Benefício Fiscal": _PERFIS[k]["beneficios"],
        }
        for k in _PERFIS if k != "Goyaço"
    ])
    st.dataframe(_df_im, use_container_width=True, hide_index=True)

    # Todos os benefícios em expanders (visão geral)
    with st.expander("📚 Ver todos os benefícios do grupo"):
        for _bid, _b in _BENEFICIOS_LEGIS.items():
            st.markdown(f"**{_b['nome']}** — {_b['uf']} | ICMS efetivo: `{_b['icms']}`")
            st.markdown(f"> {_b['resumo'][:220]}…")
            _btn_link("Ver legislação", _b["url"], "outline", "↗")
            st.markdown("---")

    st.markdown("---")
    _section("📋 Escopo do Diagnóstico Completo (em desenvolvimento)")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**Cruzamentos planejados:**
- EFD Fiscal × EFD Contribuições × ECF
- DCTF × DARF efetivamente pagos
- SPED Contábil × Balancetes
- NF-e emitidas × NF-e escrituradas
- Folha eSocial × DCTF-Web
""")
    with c2:
        st.markdown("""
**Saídas previstas:**
- Relatório de inconsistências por tributo
- Ranking de risco por obrigação
- Mapa de exposição fiscal (UF × tributo)
- Estimativa de passivo tributário
- Plano de regularização sugerido
""")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Malha Fina — Federal
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: PIS/COFINS
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "💼 PIS/COFINS":
    _header("💼", f"PIS/COFINS — {_EMP_LABELS.get(_empresa, _empresa)}",
            "Apuração, créditos e conformidade das contribuições sociais")

    st.markdown("""
    <div style="background:linear-gradient(135deg,#0D1B42,#1B3070);border-radius:12px;
                padding:24px 28px;color:white;margin-bottom:20px">
        <h3 style="color:white;margin:0 0 8px 0">🚧 Módulo em desenvolvimento</h3>
        <p style="color:rgba(255,255,255,0.75);margin:0;font-size:0.95rem">
            O módulo <strong>PIS/COFINS</strong> apurará as contribuições no regime não-cumulativo,
            identificará créditos admissíveis, divergências de alíquota e inconsistências na
            EFD Contribuições.
        </p>
    </div>
    """, unsafe_allow_html=True)

    _section("📋 Escopo previsto")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**Apuração:**
- Débito PIS/COFINS sobre receitas (CST 01/02/05/49)
- Créditos sobre insumos, energia, frete, aluguel
- Monofásico — alíquota zero e suspensão
- Substituição tributária de PIS/COFINS
- Regime cumulativo × não-cumulativo

**Divergências detectadas:**
- CST incompatível com o regime tributário
- Créditos sobre aquisições vedadas (ex: uso e consumo)
- Alíquota diferente da prevista por NCM
- Omissão de receitas tributadas × NF-e emitidas
""")
    with c2:
        st.markdown("""
**Alíquotas padrão — Não Cumulativo:**

| Tributo | Alíquota |
|---------|----------|
| PIS     | 1,65%    |
| COFINS  | 7,60%    |
| **Total** | **9,25%** |

**Alíquotas — Cumulativo (Simples/LP):**

| Tributo | Alíquota |
|---------|----------|
| PIS     | 0,65%    |
| COFINS  | 3,00%    |
| **Total** | **3,65%** |
""")

    _section("🔗 Links EFD Contribuições")
    c1, c2, c3 = st.columns(3)
    with c1:
        _btn_link("Portal SPED (RFB)", "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/acoes-e-programas/sped", "primary")
    with c2:
        _btn_link("Manual EFD Contribuições", "https://www.gov.br/receitafederal/pt-br/", "outline")
    with c3:
        _btn_link("Tabela CST PIS/COFINS", "https://www.gov.br/receitafederal/pt-br/", "outline")

    st.info("💡 Para priorizar este módulo ou contribuir com requisitos, fale com a equipe Alianzo.")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Malha Fina — Estadual
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "🕸️ Malha Fina Estadual":
    _header("🕸️", f"Malha Fina Estadual — {_EMP_LABELS.get(_empresa, _empresa)}",
            "Diagnóstico das críticas do Malha Fina: SPED × GIA × NF-e × DARE")

    st.markdown("""
    <div style="background:linear-gradient(135deg,#0D1B42,#1B3070);border-radius:12px;
                padding:24px 28px;color:white;margin-bottom:20px">
        <h3 style="color:white;margin:0 0 8px 0">🚧 Módulo em desenvolvimento</h3>
        <p style="color:rgba(255,255,255,0.75);margin:0;font-size:0.95rem">
            O módulo <strong>Malha Fina Estadual</strong> cruzará as obrigações acessórias estaduais
            (SPED EFD, GIA/SINTEGRA, NF-e) com os recolhimentos de ICMS para cada UF,
            identificando divergências antes que a SEFAZ notifique.
        </p>
    </div>
    """, unsafe_allow_html=True)

    _section("📋 Escopo previsto")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
**Cruzamentos planejados:**
- EFD ICMS/IPI × GIA por UF
- NF-e autorizadas × escrituradas no SPED
- DARE/GNRE pagos × ICMS apurado
- DIFAL recolhido × vendas interestaduais
- Substituição tributária retida × repassada
- Crédito de ICMS × vedações por produto
""")
    with c2:
        st.markdown("""
**Divergências detectadas:**
- ICMS apurado vs. recolhido (saldo em aberto)
- NF-e não escrituradas no SPED
- DIFAL insuficiente em compras interestaduais
- Estorno de crédito não realizado (CST 20)
- GIA divergente do SPED EFD entregue
- Registro D600/D695 × C100 inconsistentes
""")

    _section("🔗 Links SEFAZ / SPED")
    c1, c2, c3 = st.columns(3)
    with c1:
        _btn_link("SEFAZ-GO", "https://www.sefaz.go.gov.br/", "primary")
    with c2:
        _btn_link("Consultar NF-e emitidas", "https://www.nfe.fazenda.gov.br/portal/", "outline")
    with c3:
        _btn_link("Portal SPED", "https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/acoes-e-programas/sped", "outline")

    st.info("💡 Para priorizar este módulo ou contribuir com requisitos, fale com a equipe Alianzo.")


# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA: Agenda do Líder
# ══════════════════════════════════════════════════════════════════════════════
elif _pagina == "📅 Agenda do Líder":
    import datetime as _dt

    _header("📅", "Agenda do Líder",
            "Cronogramas, prazos e acompanhamento dos liderados")

    # ── CSS extra para agenda ─────────────────────────────────────────────────
    st.markdown("""
    <style>
    .agenda-card {
        background: white;
        border: 1px solid #CDD9F0;
        border-radius: 12px;
        padding: 18px 20px;
        margin: 8px 0;
        border-left: 5px solid #1B56A3;
        box-shadow: 0 2px 8px rgba(27,48,112,0.07);
    }
    .agenda-card.verde  { border-left-color: #1E9E5E; }
    .agenda-card.amarelo{ border-left-color: #E8A020; }
    .agenda-card.vermelho{ border-left-color: #D94040; }
    .agenda-card h4 { margin: 0 0 4px 0; color: #0D1B42; font-size: 0.97rem; }
    .agenda-card .meta { font-size: 0.8rem; color: #888; margin: 0; }
    .tag {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 700;
        margin-right: 6px;
    }
    .tag-azul    { background: #EAF1FB; color: #1B3070; }
    .tag-verde   { background: #E6F7EF; color: #1E9E5E; }
    .tag-amarelo { background: #FFF4E0; color: #b87300; }
    .tag-vermelho{ background: #FDE8E8; color: #D94040; }
    </style>
    """, unsafe_allow_html=True)

    tab_cronograma, tab_liderados, tab_tarefas = st.tabs([
        "📆 Cronograma",
        "👥 Liderados",
        "✅ Tarefas & Prazos",
    ])

    # ── ABA CRONOGRAMA ────────────────────────────────────────────────────────
    with tab_cronograma:
        _section("📆 Eventos e Prazos Fiscais — Próximos 30 dias")

        hoje = _dt.date.today()
        # ── Eventos fixos (obrigações fiscais reais de Junho–Julho/2026) ──────
        _eventos = [
            # Já realizados / hoje
            {"data": _dt.date(2026, 6, 29), "evento": "✅ Emissão CND Federal — 5 empresas (DN, EDN, R3, Cristal, Atacadão)",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 6, 29), "evento": "✅ PDFs CNDs organizados (pasta Downloads)",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 6, 29), "evento": "✅ Apuração ICMS Junho/2026 — planilha 3 abas gerada",
             "tipo": "revisao",   "responsavel": "Equipe Fiscal"},
            {"data": _dt.date(2026, 6, 29), "evento": "✅ Layout Fiscal 360 — Alianzo navy/blue implantado",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 6, 29), "evento": "✅ Módulos Certidões + Guias ICMS adicionados ao app",
             "tipo": "revisao",   "responsavel": "Éliton"},
            # Pendentes imediatos
            {"data": _dt.date(2026, 6, 30), "evento": "⏳ Push GitHub — app.py Fiscal 360 atualizado",
             "tipo": "obrigacao", "responsavel": "Éliton"},
            {"data": _dt.date(2026, 7, 3),  "evento": "Emitir Certidões Estaduais (SEFAZ-GO) — 5 empresas",
             "tipo": "obrigacao", "responsavel": "Equipe Fiscal"},
            {"data": _dt.date(2026, 7, 3),  "evento": "Emitir Certidões Municipais — Goiânia + filiais EDN",
             "tipo": "obrigacao", "responsavel": "Equipe Fiscal"},
            {"data": _dt.date(2026, 7, 5),  "evento": "Transmissão SPED Fiscal EFD ICMS/IPI — lote via PVA",
             "tipo": "obrigacao", "responsavel": "Equipe SPED"},
            {"data": _dt.date(2026, 7, 7),  "evento": "Revisão Análise de Entradas — divergências ICMS Junho",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 7, 7),  "evento": "Revisão Análise de Saídas — divergências ICMS Junho",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 7, 10), "evento": "GIA-GO — Entrega Junho/2026",
             "tipo": "obrigacao", "responsavel": "Equipe Fiscal"},
            {"data": _dt.date(2026, 7, 15), "evento": "EFD Contribuições (PIS/COFINS) — Entrega Junho/2026",
             "tipo": "obrigacao", "responsavel": "Equipe Federal"},
            {"data": _dt.date(2026, 7, 15), "evento": "EFD ICMS/IPI (SPED Fiscal) — Prazo entrega Junho",
             "tipo": "obrigacao", "responsavel": "Equipe SPED"},
            {"data": _dt.date(2026, 7, 20), "evento": "DARE-GO — Vencimento ICMS Normal Junho/2026",
             "tipo": "pagamento", "responsavel": "Financeiro"},
            {"data": _dt.date(2026, 7, 20), "evento": "DARE-GO PROTEGE/GO — Vencimento Junho/2026",
             "tipo": "pagamento", "responsavel": "Financeiro"},
            {"data": _dt.date(2026, 7, 20), "evento": "FGTS Digital — Vencimento Junho/2026",
             "tipo": "pagamento", "responsavel": "Financeiro"},
            {"data": _dt.date(2026, 7, 20), "evento": "DARE-GO DIFAL — Vencimento Junho/2026",
             "tipo": "pagamento", "responsavel": "Financeiro"},
            {"data": _dt.date(2026, 7, 20), "evento": "DCTF — Entrega Junho/2026",
             "tipo": "obrigacao", "responsavel": "Equipe Federal"},
            {"data": _dt.date(2026, 7, 25), "evento": "Check-in mensal com liderados — Julho/2026",
             "tipo": "reuniao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 7, 28), "evento": "Diagnóstico Malha Fina Estadual — EDN / Atacadão / Cristal",
             "tipo": "revisao",   "responsavel": "Éliton"},
            {"data": _dt.date(2026, 8, 5),  "evento": "Imersão Fiscal — Relatório diagnóstico completo multitributo",
             "tipo": "revisao",   "responsavel": "Éliton"},
        ]
        _eventos.sort(key=lambda x: x["data"])

        _cores = {"obrigacao": "🔴", "pagamento": "🟡", "reuniao": "🔵", "revisao": "🟢"}
        _tags  = {"obrigacao": "Obrigação", "pagamento": "Pagamento", "reuniao": "Reunião", "revisao": "Revisão"}
        _tag_css = {"obrigacao": "tag-vermelho", "pagamento": "tag-amarelo", "reuniao": "tag-azul", "revisao": "tag-verde"}

        for ev in _eventos:
            dias = (ev["data"] - hoje).days
            _urgencia = "vermelho" if dias <= 3 else ("amarelo" if dias <= 7 else "verde")
            st.markdown(f"""
            <div class="agenda-card {_urgencia}">
                <h4>{_cores[ev["tipo"]]} {ev["evento"]}</h4>
                <p class="meta">
                    <span class="tag {_tag_css[ev['tipo']]}">{_tags[ev['tipo']]}</span>
                    📅 {ev['data'].strftime('%d/%m/%Y')} &nbsp;·&nbsp;
                    ⏳ {'Hoje!' if dias == 0 else f'{dias} dias'} &nbsp;·&nbsp;
                    👤 {ev['responsavel']}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        _section("➕ Adicionar Evento ao Cronograma")
        with st.form("form_evento"):
            fe1, fe2, fe3 = st.columns([2, 1, 1])
            with fe1:
                _nome_ev  = st.text_input("Descrição do evento")
            with fe2:
                _data_ev  = st.date_input("Data", value=hoje + _dt.timedelta(days=7))
            with fe3:
                _tipo_ev  = st.selectbox("Tipo", ["obrigacao","pagamento","reuniao","revisao"],
                                          format_func=lambda x: _tags[x])
            _resp_ev = st.text_input("Responsável")
            if st.form_submit_button("Salvar evento", type="primary"):
                if "agenda_eventos" not in st.session_state:
                    st.session_state["agenda_eventos"] = []
                st.session_state["agenda_eventos"].append({
                    "data": _data_ev, "evento": _nome_ev,
                    "tipo": _tipo_ev, "responsavel": _resp_ev,
                })
                st.success(f"✅ Evento '{_nome_ev}' salvo para {_data_ev.strftime('%d/%m/%Y')}.")

        # Eventos adicionados na sessão
        if st.session_state.get("agenda_eventos"):
            _section("📌 Eventos adicionados nesta sessão")
            for ev in sorted(st.session_state["agenda_eventos"], key=lambda x: x["data"]):
                dias = (ev["data"] - hoje).days
                st.markdown(f"""
                <div class="agenda-card">
                    <h4>{_cores.get(ev['tipo'],'📌')} {ev['evento']}</h4>
                    <p class="meta">
                        <span class="tag tag-azul">{_tags.get(ev['tipo'],'')}</span>
                        📅 {ev['data'].strftime('%d/%m/%Y')} &nbsp;·&nbsp;
                        ⏳ {'Hoje!' if dias == 0 else (f'{dias} dias' if dias >= 0 else f'{abs(dias)} dias atrás')} &nbsp;·&nbsp;
                        👤 {ev['responsavel']}
                    </p>
                </div>
                """, unsafe_allow_html=True)

    # ── ABA LIDERADOS ─────────────────────────────────────────────────────────
    with tab_liderados:
        _section("👥 Cadastro de Liderados")

        if "liderados" not in st.session_state:
            st.session_state["liderados"] = [
                {"nome": "Liderado 1", "cargo": "Analista Fiscal",   "area": "ICMS",         "status": "✅ Em dia"},
                {"nome": "Liderado 2", "cargo": "Analista Tributário","area": "PIS/COFINS",   "status": "⚠️ Pendência"},
                {"nome": "Liderado 3", "cargo": "Assistente Fiscal",  "area": "SPED/Certidões","status": "✅ Em dia"},
            ]

        import pandas as _pd_lid
        _df_lid = _pd_lid.DataFrame(st.session_state["liderados"])
        st.dataframe(_df_lid, use_container_width=True, hide_index=True)

        st.markdown("---")
        with st.expander("➕ Adicionar / editar liderado"):
            with st.form("form_liderado"):
                l1, l2, l3, l4 = st.columns(4)
                _l_nome   = l1.text_input("Nome")
                _l_cargo  = l2.text_input("Cargo")
                _l_area   = l3.text_input("Área")
                _l_status = l4.selectbox("Status", ["✅ Em dia","⚠️ Pendência","🔴 Atrasado","🔵 Em férias"])
                if st.form_submit_button("Adicionar", type="primary"):
                    st.session_state["liderados"].append({
                        "nome": _l_nome, "cargo": _l_cargo,
                        "area": _l_area, "status": _l_status,
                    })
                    st.rerun()

    # ── ABA TAREFAS ───────────────────────────────────────────────────────────
    with tab_tarefas:
        _section("✅ Tarefas & Prazos por Liderado")

        if "tarefas_agenda" not in st.session_state:
            st.session_state["tarefas_agenda"] = [
                # ── Concluídas ────────────────────────────────────────────────
                {"tarefa": "Emitir CND Federal — DN Armazenamento (28.221.185/0001-83)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Emitir CND Federal — EDN Utilidades (20.758.851/0001-05)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Emitir CND Federal — R3 Suprimentos (10.641.901/0001-16)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Emitir CND Federal — Comercial Cristal (07.515.610/0001-77)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Emitir CND Federal — Atacadão do Lar (35.917.755/0001-30)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Organizar PDFs CNDs em pasta (5 arquivos nomeados por empresa/CNPJ)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Gerar planilha Apuração ICMS Junho/2026 — 3 abas (Apuração, Entradas, Saídas)",
                 "responsavel": "Equipe Fiscal", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Implantar layout Alianzo (navy #1B3070 / blue #1B56A3) no Fiscal 360",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Adicionar módulo Certidões (Federal / Estadual / Municipal) ao app",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Adicionar módulo Guias ICMS (DARE-GO, DIFAL, Calendário) ao app",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Renomear plataforma para Fiscal 360 (remover referências ICMS/GO)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                {"tarefa": "Criar módulos Imersão Fiscal, Malha Fina, PIS/COFINS, Agenda do Líder",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 29), "status": "Concluído"},
                # ── Pendentes imediatas ────────────────────────────────────────
                {"tarefa": "Push GitHub — app.py Fiscal 360 (versão atual)",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 6, 30), "status": "Pendente"},
                {"tarefa": "Emitir Certidões Estaduais SEFAZ-GO — 5 empresas",
                 "responsavel": "Equipe Fiscal", "prazo": _dt.date(2026, 7, 3), "status": "Pendente"},
                {"tarefa": "Emitir Certidões Municipais — Goiânia + filiais EDN",
                 "responsavel": "Equipe Fiscal", "prazo": _dt.date(2026, 7, 3), "status": "Pendente"},
                {"tarefa": "Transmissão SPED Fiscal EFD ICMS/IPI — lote via PVA (automação)",
                 "responsavel": "Equipe SPED", "prazo": _dt.date(2026, 7, 5), "status": "Pendente"},
                # ── Em andamento ──────────────────────────────────────────────
                {"tarefa": "Análise divergências ICMS Entradas — Junho/2026",
                 "responsavel": "Equipe Fiscal", "prazo": _dt.date(2026, 7, 7), "status": "Em andamento"},
                {"tarefa": "Análise divergências ICMS Saídas — Junho/2026",
                 "responsavel": "Equipe Fiscal", "prazo": _dt.date(2026, 7, 7), "status": "Em andamento"},
                {"tarefa": "Diagnóstico Malha Fina Estadual — EDN / Atacadão / Cristal",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 7, 28), "status": "Em andamento"},
                {"tarefa": "Apuração PIS/COFINS — desenvolvimento do módulo no Fiscal 360",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 8, 1), "status": "Em andamento"},
                {"tarefa": "Imersão Fiscal — relatório diagnóstico completo multitributo",
                 "responsavel": "Éliton", "prazo": _dt.date(2026, 8, 5), "status": "Em andamento"},
            ]

        _status_icon = {"Concluído": "✅", "Em andamento": "🔄", "Pendente": "⏳", "Bloqueado": "🔴"}
        _status_cor  = {"Concluído": "verde", "Em andamento": "amarelo", "Pendente": "amarelo", "Bloqueado": "vermelho"}

        for i, t in enumerate(st.session_state["tarefas_agenda"]):
            dias = (t["prazo"] - hoje).days
            _prazo_txt = "✅ Concluído" if t["status"] == "Concluído" else (
                "🔴 Vencido!" if dias < 0 else ("⚠️ Hoje!" if dias == 0 else f"⏳ {dias} dias"))
            st.markdown(f"""
            <div class="agenda-card {_status_cor.get(t['status'],'verde')}">
                <h4>{_status_icon.get(t['status'],'📌')} {t['tarefa']}</h4>
                <p class="meta">
                    👤 {t['responsavel']} &nbsp;·&nbsp;
                    📅 {t['prazo'].strftime('%d/%m/%Y')} &nbsp;·&nbsp;
                    {_prazo_txt}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")
        with st.expander("➕ Nova tarefa"):
            with st.form("form_tarefa"):
                t1, t2 = st.columns(2)
                _t_nome  = t1.text_input("Descrição da tarefa")
                _t_resp  = t2.text_input("Responsável")
                t3, t4 = st.columns(2)
                _t_prazo = t3.date_input("Prazo", value=hoje + _dt.timedelta(days=7))
                _t_stat  = t4.selectbox("Status", ["Pendente","Em andamento","Concluído","Bloqueado"])
                if st.form_submit_button("Adicionar tarefa", type="primary"):
                    st.session_state["tarefas_agenda"].append({
                        "tarefa": _t_nome, "responsavel": _t_resp,
                        "prazo": _t_prazo, "status": _t_stat,
                    })
                    st.rerun()


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="alz-footer">
    Alianzo Fiscal 360 &nbsp;·&nbsp; Plataforma Tributária Multitributo &nbsp;·&nbsp;
    Dados processados em memória — não armazenados no servidor &nbsp;·&nbsp; 2026
</div>
""", unsafe_allow_html=True)
