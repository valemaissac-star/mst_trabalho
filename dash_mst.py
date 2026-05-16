import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from collections import Counter
import os
from pathlib import Path

st.set_page_config(
    page_title="Propriedade Digital em Jogos - Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ======================================================
# CONFIGURAÇÃO DE CAMINHOS
# ======================================================

# Obter diretório base (onde está o .py)
BASE_DIR = Path(__file__).parent
ARQUIVOS_DIR = BASE_DIR / "arquivos"

# Possíveis localizações do arquivo
POSSIBLE_PATHS = [
    ARQUIVOS_DIR / "formulario_expandido.xlsx",
    BASE_DIR / "formulario_expandido.xlsx",
    Path("arquivos") / "formulario_expandido.xlsx",
    Path("formulario_expandido.xlsx"),
]

# ======================================================
# TEMA VISUAL GLOBAL
# ======================================================

st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    }

    h1 { font-size: 2.8rem !important; font-weight: 900 !important; color: #ffffff !important; }
    h2 { font-size: 2rem !important; font-weight: 800 !important; color: #e0e0ff !important; }
    h3 { font-size: 1.6rem !important; font-weight: 700 !important; color: #c0c0ff !important; }

    p, div, span, label { font-size: 1.05rem !important; }

    [data-testid="metric-container"] {
        background: linear-gradient(135deg, #1a1a3e, #2d2d6b);
        border: 1px solid #4444aa;
        border-radius: 16px;
        padding: 20px !important;
        box-shadow: 0 0 20px rgba(100, 100, 255, 0.3);
    }
    [data-testid="metric-container"] label {
        font-size: 1rem !important;
        color: #aaaaff !important;
        font-weight: 700 !important;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 900 !important;
        color: #ffffff !important;
    }

    .stSelectbox label { font-size: 1.1rem !important; font-weight: 700 !important; color: #aaaaff !important; }
    .stDataFrame { border-radius: 12px; overflow: hidden; }
    hr { border-color: #4444aa !important; opacity: 0.4; }

    .stDownloadButton button {
        background: linear-gradient(135deg, #6600cc, #9900ff) !important;
        color: white !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        font-size: 1rem !important;
        border: none !important;
    }

    .metric-principal {
        background: linear-gradient(135deg, #2d1b69, #4a2c8e);
        border: 2px solid #7c5cdb;
        border-radius: 20px;
        padding: 30px !important;
        box-shadow: 0 0 30px rgba(124, 92, 219, 0.4);
        text-align: center;
    }
    .metric-principal-value {
        font-size: 3.5rem !important;
        font-weight: 900 !important;
        color: #FFE66D !important;
        margin: 10px 0;
    }
    .metric-principal-label {
        font-size: 1.3rem !important;
        color: #c0c0ff !important;
        font-weight: 700 !important;
    }

    .stTabs [data-baseweb="tab-list"] {
        background: linear-gradient(135deg, #1a1a3e, #2d2d6b);
        border-radius: 14px;
        padding: 6px;
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 10px;
        color: #aaaaff !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6600cc, #9900ff) !important;
        color: white !important;
    }

    .card-insight {
        background: linear-gradient(135deg, #1a1a3e, #2d2d6b);
        border-left: 4px solid #A855F7;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

CORES_VIBRANTES = [
    "#FF6B6B", "#FFE66D", "#4ECDC4", "#A855F7",
    "#F97316", "#22D3EE", "#10B981", "#F43F5E",
    "#6366F1", "#84CC16", "#EC4899", "#14B8A6",
    "#F59E0B", "#8B5CF6", "#06B6D4", "#EF4444"
]

TEMPLATE_GRAFICO = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(15,12,41,0.0)",
    plot_bgcolor="rgba(15,12,41,0.0)",
    font=dict(family="Inter, sans-serif", size=14, color="#e0e0ff"),
    title_font=dict(size=20, color="#ffffff", family="Inter, sans-serif"),
)

# ======================================================
# CARREGAMENTO DOS DADOS (OTIMIZADO)
# ======================================================

@st.cache_data
def load_data():
    """
    Carrega dados do Excel com múltiplos caminhos possíveis.
    Tenta carregar de vários locais até encontrar o arquivo.
    """
    
    # Mostrar caminhos sendo testados (debug)
    st.sidebar.write("🔍 Testando caminhos...")
    
    for path in POSSIBLE_PATHS:
        try:
            if path.exists():
                st.sidebar.success(f"✅ Arquivo encontrado em: {path}")
                df = pd.read_excel(path)
                st.sidebar.info(f"📊 {len(df)} respostas carregadas!")
                return df
        except Exception as e:
            st.sidebar.write(f"❌ Erro em {path}: {str(e)[:50]}")
            continue
    
    # Se nenhum caminho funcionou, mostrar erro detalhado
    erro_msg = f"""
    ❌ ARQUIVO NÃO ENCONTRADO!
    
    Locais testados:
    {chr(10).join([f'  • {p}' for p in POSSIBLE_PATHS])}
    
    Arquivos disponíveis em '{ARQUIVOS_DIR}':{chr(10)}{str(list(ARQUIVOS_DIR.glob('*')) if ARQUIVOS_DIR.exists() else 'Pasta não existe')}
    """
    raise FileNotFoundError(erro_msg)

# ======================================================
# TENTAR CARREGAR DADOS
# ======================================================

try:
    df_raw = load_data()
except FileNotFoundError as e:
    st.error(str(e))
    st.info("""
    📋 **SOLUÇÃO:**
    1. Verifique se `formulario_expandido.xlsx` está em `/arquivos/`
    2. Ou coloque na mesma pasta do `dash_mst.py`
    3. Reinicie o dashboard
    """)
    st.stop()

# ======================================================
# PROCESSAMENTO BASE
# ======================================================

df = df_raw.copy()
df.columns = df.columns.str.strip()

# Renomear para mais fácil acesso (ajustar conforme suas colunas)
col_frequencia = [c for c in df.columns if 'frequência' in c.lower()][0] if any('frequência' in c.lower() for c in df.columns) else None
col_acesso = [c for c in df.columns if 'geralmente acessa' in c.lower()][0] if any('geralmente acessa' in c.lower() for c in df.columns) else None
col_midia = [c for c in df.columns if 'mídia física' in c.lower()][0] if any('mídia física' in c.lower() for c in df.columns) else None
col_plataforma = [c for c in df.columns if 'plataforma' in c.lower() and 'mais costuma' in c.lower()][0] if any('plataforma' in c.lower() and 'mais costuma' in c.lower() for c in df.columns) else None
col_perda = [c for c in df.columns if 'perdeu acesso' in c.lower()][0] if any('perdeu acesso' in c.lower() for c in df.columns) else None
col_sentimento = [c for c in df.columns if 'perdesse acesso' in c.lower() and 'como se sentiria' in c.lower()][0] if any('perdesse acesso' in c.lower() and 'como se sentiria' in c.lower() for c in df.columns) else None
col_propriedade = [c for c in df.columns if 'realmente seu' in c.lower()][0] if any('realmente seu' in c.lower() for c in df.columns) else None
col_conhecimento = [c for c in df.columns if 'licenciados' in c.lower()][0] if any('licenciados' in c.lower() for c in df.columns) else None
col_restricoes = [c for c in df.columns if 'problemas com restrições' in c.lower()][0] if any('problemas com restrições' in c.lower() for c in df.columns) else None
col_seguranca = [c for c in df.columns if 'seguro quanto' in c.lower()][0] if any('seguro quanto' in c.lower() for c in df.columns) else None
col_opiniao = [c for c in df.columns if 'justo que' in c.lower() and 'Por quê' in c][0] if any('justo que' in c.lower() and 'Por quê' in c for c in df.columns) else None

# Criar aliases mais limpos (apenas se a coluna existe)
if col_frequencia:
    df['frequencia_consumo'] = df[col_frequencia].str.lower().str.strip()
if col_acesso:
    df['forma_acesso'] = df[col_acesso].str.lower().str.strip()
if col_midia:
    df['midia_fisica'] = df[col_midia].str.lower().str.strip()
if col_plataforma:
    df['plataforma'] = df[col_plataforma].str.lower().str.strip()
if col_perda:
    df['perda_acesso'] = df[col_perda].str.lower().str.strip()
if col_sentimento:
    df['sentimento_perda'] = df[col_sentimento].str.lower().str.strip()
if col_propriedade:
    df['sente_propriedade'] = df[col_propriedade].str.lower().str.strip()
if col_conhecimento:
    df['conhece_licenca'] = df[col_conhecimento].str.lower().str.strip()
if col_restricoes:
    df['problemas_restricoes'] = df[col_restricoes].str.lower().str.strip()
if col_seguranca:
    df['seguranca_biblioteca'] = df[col_seguranca].str.lower().str.strip()
if col_opiniao:
    df['opiniao_justica'] = df[col_opiniao].str.lower().str.strip()

# Converter timestamp
if 'Carimbo de data/hora' in df.columns:
    df['data'] = pd.to_datetime(df['Carimbo de data/hora'], errors='coerce')
else:
    df['data'] = pd.Timestamp.now()

# ======================================================
# FUNÇÕES AUXILIARES
# ======================================================

def calcular_metricas_gerais():
    return {
        'total_respostas': len(df),
        'taxa_perda': (df['perda_acesso'] != 'não, nunca').sum() / len(df) * 100 if 'perda_acesso' in df.columns else 0,
        'taxa_problemas': (df['problemas_restricoes'] == 'sim').sum() / len(df) * 100 if 'problemas_restricoes' in df.columns else 0,
        'taxa_conhece_licenca': (df['conhece_licenca'] == 'sim').sum() / len(df) * 100 if 'conhece_licenca' in df.columns else 0,
    }

# ======================================================
# HEADER
# ======================================================

st.markdown("""
<div style='text-align:center; padding: 2rem 0 1rem 0;'>
    <h1> TRABALHO MST</h1>
    <p style='color:#aaaaff; font-size:1.2rem !important;'>ANÁLISE DE DADOS</p>
</div>
""", unsafe_allow_html=True)
st.markdown("---")

# ======================================================
# ABAS PRINCIPAIS
# ======================================================

aba_visao_geral, aba_dados = st.tabs([
    "📊 Visão Geral",
    "📋 Dados Brutos"
])

# ══════════════════════════════════════════════════════
# ABA 1 — VISÃO GERAL
# ══════════════════════════════════════════════════════
with aba_visao_geral:

    st.markdown("## 📊 Visão Geral da Pesquisa")

    metricas = calcular_metricas_gerais()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""<div class='metric-principal'>
            <div class='metric-principal-label'>Total de Respostas</div>
            <div class='metric-principal-value'>{metricas['total_respostas']}</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown(f"""<div class='metric-principal'>
            <div class='metric-principal-label'>Já Perderam Acesso</div>
            <div class='metric-principal-value'>{metricas['taxa_perda']:.1f}%</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown(f"""<div class='metric-principal'>
            <div class='metric-principal-label'>Conhecem Licenças</div>
            <div class='metric-principal-value'>{metricas['taxa_conhece_licenca']:.1f}%</div>
        </div>""", unsafe_allow_html=True)
    with col4:
        st.markdown(f"""<div class='metric-principal'>
            <div class='metric-principal-label'>Com Problemas</div>
            <div class='metric-principal-value'>{metricas['taxa_problemas']:.1f}%</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("---")

    # Distribuições por coluna disponível
    if 'frequencia_consumo' in df.columns:
        st.markdown("### 🎮 Frequência de Consumo de Jogos")
        freq_counts = df['frequencia_consumo'].value_counts()
        
        fig_freq = go.Figure(data=[go.Bar(
            x=freq_counts.index,
            y=freq_counts.values,
            text=freq_counts.values,
            textposition='outside',
            marker=dict(
                color=freq_counts.values,
                colorscale="Viridis",
                showscale=False,
                line=dict(color="rgba(255,255,255,0.15)", width=1)
            ),
        )])
        fig_freq.update_layout(
            **TEMPLATE_GRAFICO,
            title="📊 Frequência de Consumo",
            height=400,
            xaxis=dict(title="Frequência"),
            yaxis=dict(title="Quantidade"),
        )
        st.plotly_chart(fig_freq, use_container_width=True, key="freq_consumo")

        st.markdown("---")

    # Gráficos de outras colunas
    col_ac1, col_ac2 = st.columns(2)

    with col_ac1:
        if 'forma_acesso' in df.columns:
            acesso_counts = df['forma_acesso'].value_counts()
            fig_acesso = go.Figure(data=[go.Pie(
                labels=acesso_counts.index,
                values=acesso_counts.values,
                marker=dict(colors=CORES_VIBRANTES, line=dict(color="#0f0c29", width=3)),
                textfont=dict(size=14),
                hole=0.4
            )])
            fig_acesso.update_layout(**TEMPLATE_GRAFICO, title="💳 Como Acessa Jogos", height=420)
            st.plotly_chart(fig_acesso, use_container_width=True, key="forma_acesso")

    with col_ac2:
        if 'plataforma' in df.columns:
            plataforma_counts = df['plataforma'].value_counts()
            fig_plat = go.Figure(data=[go.Pie(
                labels=plataforma_counts.index,
                values=plataforma_counts.values,
                marker=dict(colors=CORES_VIBRANTES, line=dict(color="#0f0c29", width=3)),
                textfont=dict(size=14),
                hole=0.4
            )])
            fig_plat.update_layout(**TEMPLATE_GRAFICO, title="🖥️ Plataformas", height=420)
            st.plotly_chart(fig_plat, use_container_width=True, key="plataforma")

    st.markdown("---")

    # Propriedade e Segurança
    col_prop1, col_prop2 = st.columns(2)

    with col_prop1:
        if 'sente_propriedade' in df.columns:
            prop_counts = df['sente_propriedade'].value_counts()
            fig_prop = go.Figure(data=[go.Bar(
                y=prop_counts.index,
                x=prop_counts.values,
                orientation='h',
                text=prop_counts.values,
                textposition='outside',
                marker=dict(color=prop_counts.values, colorscale="RdYlGn_r", showscale=False),
            )])
            fig_prop.update_layout(
                **TEMPLATE_GRAFICO,
                title="🎮 Sente que o Jogo é Realmente Seu?",
                height=400,
                xaxis=dict(title="Quantidade"),
            )
            st.plotly_chart(fig_prop, use_container_width=True, key="sente_propriedade")

    with col_prop2:
        if 'seguranca_biblioteca' in df.columns:
            seg_counts = df['seguranca_biblioteca'].value_counts()
            fig_seg = go.Figure(data=[go.Bar(
                y=seg_counts.index,
                x=seg_counts.values,
                orientation='h',
                text=seg_counts.values,
                textposition='outside',
                marker=dict(color=seg_counts.values, colorscale="RdYlGn", showscale=False),
            )])
            fig_seg.update_layout(
                **TEMPLATE_GRAFICO,
                title="🛡️ Segurança na Biblioteca Digital",
                height=400,
                xaxis=dict(title="Quantidade"),
            )
            st.plotly_chart(fig_seg, use_container_width=True, key="seguranca")

# ══════════════════════════════════════════════════════
# ABA 2 — DADOS BRUTOS
# ══════════════════════════════════════════════════════
with aba_dados:

    st.markdown("## 📋 Dados Brutos")

    # Tabela completa
    st.markdown("### 📊 Todas as Respostas")
    st.dataframe(df, use_container_width=True)

    st.markdown("---")

    # Exportar
    st.markdown("### 💾 Exportar Dados")
    
    st.download_button(
        "📥 Exportar em CSV",
        df.to_csv(index=False, encoding='utf-8-sig'),
        "analise_propriedade_digital.csv",
        "text/csv"
    )

st.markdown("---")
st.markdown("<p style='text-align:center;color:#6666aa;font-size:0.95rem;'>Dashboard - Propriedade Digital em Jogos · IFCE Campus Aracati</p>",
            unsafe_allow_html=True)