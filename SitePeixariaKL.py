import os
import unicodedata
import urllib.parse
import pandas as pd
import streamlit as st

# ==========================================
# CONFIGURAÇÕES DA PEIXARIA KL
# ==========================================
NOME_LOJA = "Peixaria KL"
WHATSAPP = "5516997179478"
TELEFONE_EXIBICAO = "(16) 99717-9478"
SLOGAN = "Peixes frescos, frutos do mar e porções selecionadas direto para sua mesa"
HORARIO_ATENDIMENTO = "Seg a Sáb: 08h às 18h • Dom: 08h às 12h"
ENDERECO_LOJA = "Rua Sérgio Felício de Souza, 995 - Quarto Centenário, Matão - SP"
CIDADE_REGIAO = "Matão e Região - SP"
TAXA_ENTREGA = 5.00
LINK_PLANILHA_CSV = ""

PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL_LOCAL = os.path.join(PASTA_ATUAL, "Peixaria KL.xlsx")

st.set_page_config(
    page_title=f"{NOME_LOJA} | Catálogo Digital",
    page_icon="🐟",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "modo_escuro" not in st.session_state:
    st.session_state.modo_escuro = False

if "carrinho" not in st.session_state:
    st.session_state.carrinho = {}

if "categoria_ativa" not in st.session_state:
    st.session_state.categoria_ativa = "Todas"

# ==========================================
# PALETA DINÂMICA (CLARO / ESCURO)
# ==========================================
if st.session_state.modo_escuro:
    bg_app = "linear-gradient(180deg, #091326 0%, #0f172a 50%, #020617 100%)"
    bg_card = "#1e293b"
    border_card = "#334155"
    text_main = "#f8fafc"
    text_muted = "#94a3b8"
    price_color = "#38bdf8"
    bg_cat_badge = "#0c4a6e"
    text_cat_badge = "#7dd3fc"
    border_cat_badge = "#0369a1"
    bg_weight_badge = "#0f172a"
    text_weight_badge = "#cbd5e1"
    border_weight_badge = "#334155"
    border_divider = "#334155"
else:
    bg_app = "linear-gradient(180deg, #eef6fc 0%, #f8fafc 40%, #f1f5f9 100%)"
    bg_card = "#ffffff"
    border_card = "#e2e8f0"
    text_main = "#0f172a"
    text_muted = "#64748b"
    price_color = "#0369a1"
    bg_cat_badge = "#f0f9ff"
    text_cat_badge = "#0284c7"
    border_cat_badge = "#bae6fd"
    bg_weight_badge = "#f8fafc"
    text_weight_badge = "#475569"
    border_weight_badge = "#e2e8f0"
    border_divider = "#f1f5f9"

# ==========================================
# DESIGN MODERNO (CSS CUSTOMIZADO)
# ==========================================
st.markdown(
    f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: 'Plus Jakarta Sans', sans-serif;
}}
.stApp {{
    background: {bg_app};
}}

/* Esconde menu padrão superior do Streamlit */
#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
header[data-testid="stHeader"] {{background: transparent;}}

/* Banner Hero Moderno */
.hero-modern {{
    background: radial-gradient(circle at top right, #0ea5e9 0%, #0369a1 35%, #081b33 100%);
    padding: 38px 28px;
    border-radius: 24px;
    color: white;
    text-align: center;
    box-shadow: 0 20px 40px -15px rgba(8, 27, 51, 0.35);
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
}}
.hero-pill {{
    display: inline-block;
    background: rgba(255, 255, 255, 0.14);
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255, 255, 255, 0.25);
    color: #7dd3fc;
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 6px 16px;
    border-radius: 999px;
    margin-bottom: 12px;
}}
.stApp .hero-modern h1 {{
    color: #ffffff !important;
    margin: 0;
    font-size: 2.9rem;
    font-weight: 800;
    letter-spacing: -1px;
    line-height: 1.1;
}}
.stApp .hero-modern h1 span {{
    background: linear-gradient(90deg, #38bdf8, #67e8f9);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}}
.stApp .hero-modern p {{
    margin: 12px auto 22px auto;
    max-width: 650px;
    font-size: 1.05rem;
    color: #e0f2fe !important;
    font-weight: 400;
}}
.hero-features {{
    display: flex;
    justify-content: center;
    flex-wrap: wrap;
    gap: 10px;
}}
.feature-chip {{
    background: rgba(8, 27, 51, 0.45);
    border: 1px solid rgba(125, 211, 252, 0.3);
    padding: 7px 15px;
    border-radius: 12px;
    font-size: 0.84rem;
    color: #f0f9ff !important;
    font-weight: 600;
}}

/* Card de Produto Moderno */
.product-card {{
    background: {bg_card};
    border-radius: 18px;
    padding: 18px;
    border: 1px solid {border_card};
    box-shadow: 0 4px 16px rgba(15, 23, 42, 0.08);
    transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    margin-bottom: 10px;
}}
.product-card:hover {{
    transform: translateY(-4px);
    box-shadow: 0 14px 28px -8px rgba(2, 132, 199, 0.25);
    border-color: #38bdf8;
}}
.card-top-row {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 12px;
}}
.cat-badge {{
    background: {bg_cat_badge};
    color: {text_cat_badge};
    font-size: 0.74rem;
    font-weight: 700;
    padding: 4px 11px;
    border-radius: 999px;
    border: 1px solid {border_cat_badge};
}}
.weight-badge {{
    background: {bg_weight_badge};
    color: {text_weight_badge};
    font-size: 0.76rem;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 8px;
    border: 1px solid {border_weight_badge};
}}

/* Foto limpa e responsiva quando cadastrada na planilha */
.product-photo-box {{
    width: 100%;
    height: 170px;
    border-radius: 14px;
    overflow: hidden;
    margin-bottom: 12px;
    background-color: #0f172a;
}}
.product-photo-img {{
    width: 100%;
    height: 100%;
    object-fit: cover;
    display: block;
    transition: transform 0.3s ease;
}}
.product-card:hover .product-photo-img {{
    transform: scale(1.05);
}}

.prod-name {{
    color: {text_main};
    font-size: 1.12rem;
    font-weight: 700;
    line-height: 1.35;
    min-height: 48px;
    margin: 6px 0 14px 0;
    text-align: left;
}}
.price-row {{
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    padding-top: 12px;
    border-top: 1px solid {border_divider};
}}
.price-label {{
    font-size: 0.78rem;
    color: {text_muted};
    font-weight: 500;
}}
.price-number {{
    font-size: 1.45rem;
    font-weight: 800;
    color: {price_color};
    letter-spacing: -0.5px;
}}

/* Botões Streamlit Estilizados */
div.stButton > button {{
    border-radius: 12px;
    font-weight: 600;
    transition: all 0.2s ease;
}}
div.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
    border: none;
    color: white;
    box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25);
}}

/* Rodapé Moderno */
.modern-footer {{
    background: #081b33;
    color: #cbd5e1;
    padding: 34px 24px;
    border-radius: 22px;
    text-align: center;
    margin-top: 48px;
    box-shadow: 0 -4px 20px rgba(8, 27, 51, 0.08);
}}
.stApp .modern-footer h3 {{
    color: #ffffff !important;
    font-weight: 800;
    margin-bottom: 8px;
}}

/* Textos gerais acompanhando o tema escolhido */
.stApp h1, .stApp h2, .stApp h3, .stApp label, .stApp .stMarkdown p, .stApp [data-testid="stCaptionContainer"] {{
    color: {text_main} !important;
}}
.stApp .hero-modern h1, .stApp .hero-modern p, .stApp .modern-footer h3, .stApp .modern-footer p, section[data-testid="stSidebar"] * {{
    color: #ffffff !important;
}}
</style>""",
    unsafe_allow_html=True,
)

ICONES_CATEGORIA = {
    "Filés": "🐟",
    "Postas": "🥩",
    "Peixes Inteiros": "🐠",
    "Camarões": "🦐",
    "Prontos para Fritar": "🍤",
    "Diversos": "🦀",
}

FOTOS_ESPECIFICAS = {
    # --- FILÉS ---
    "Filé de Bacalhau": "",
    "Filé de Cavalinha": "",
    "Filé de Corvina": "",
    "Filé de Merluza": "",
    "Filézinho de Pacu": "",
    "Filé de Pangasius": "",
    "Filé de Pirarucu em Pedaços": "",
    "Filé de Porquinho": "",
    "Filé de Salmão": "",
    "Filé de Tilápia": "",
    "Filé de Tucunaré": "",
    "Filé de Trilha": "",
    "Filé de Uritinga (Pintado da Amazônia)": "",
    "Sardinha Espalmada": "",
    "Traíra Espalmada Retalhada": "",
    # --- POSTAS ---
    "Costelinha de Pacu": "",
    "Posta de Barbado": "",
    "Posta de Cascudo": "",
    "Posta de Cação": "",
    "Posta de Tilápia": "",
    "Posta de Pintado": "",
    # --- PEIXES INTEIROS ---
    "Pacu Dessosado": "",
    "Corvina (G)": "",
    "Curimba": "",
    "Lambari": "",
    "Mandi": "",
    "Manjubinha Eviscerada sem Cabeça": "",
    "Pacu / Tambaqui com Escamas": "",
    "Piapara": "",
    "Piau 3 Pintas": "",
    "Sardinha Inteira Eviscerada sem Cabeça": "",
    "Tainha": "",
    "Toquinho de Cascudo": "",
    "Tilápia Eviscerada com Escamas": "",
    "Tilápia Dessosada sem Escamas": "",
    "Traíra": "",
    "Traíra (G)": "",
    "Tucunaré (G)": "",
    # --- CAMARÕES ---
    "Camarão (P)": "",
    "Camarão (M)": "",
    "Camarão (G)": "",
    # --- PRONTOS PARA FRITAR ---
    "Bolinho de Bacalhau": "",
    "Bolinho de Bacalhau Gourmet": "",
    "Bolinho de Camarão": "",
    "Bolinho de Carne Seca com Mandioca": "",
    "Bolinho de Costela": "",
    "Bolinho de Salmão": "",
    "Bolinho de Tilápia com Queijo": "",
    "Camarão Empanado com Requeijão": "",
    "Camarão Empanado": "",
    "Chips de Tilápia Empanado": "",
    "Espetinho de Tilápia com Bacon": "",
    "Espetinho de Tilápia com Provolone": "",
    "Filé de Porquinho Empanado": "",
    "Iscas de Pintado Empanado": "",
    "Lambari Desossado e Empanado": "",
    "Moldado de Tilápia Empanado": "",
    "Queijo Cura Empanado": "",
    "Tiras de Merluza Empanada": "",
    "Tirinhas de Pangasius Empanadas": "",
    "Tirinhas de Tilápia Empanadas": "",
    "Torresmo de Pacu Empanado": "",
    "Iscas de Filé de Frango Empanado": "",
    # --- DIVERSOS ---
    "Anel de Lula": "",
    "Caranguejo": "",
    "Farinha para Empanar": "",
    "Kani Kama": "",
    "Mandioca": "",
    "Mexilhão Desconchado": "",
    "Sardinha Salgada": "",
    "Siri": "",
}

DADOS_PADRAO = [
    ("Filés", "Filé de Bacalhau", "kg", 57.00, "Sim", ""),
    ("Filés", "Filé de Cavalinha", "500g", 15.00, "Sim", ""),
    ("Filés", "Filé de Corvina", "500g", 19.00, "Sim", ""),
    ("Filés", "Filé de Corvina", "1kg", 37.00, "Sim", ""),
    ("Filés", "Filé de Merluza", "500g", 19.50, "Sim", ""),
    ("Filés", "Filézinho de Pacu", "500g", 21.00, "Sim", ""),
    ("Filés", "Filé de Pangasius", "1kg", 40.00, "Sim", ""),
    ("Filés", "Filé de Pirarucu em Pedaços", "500g", 26.00, "Sim", ""),
    ("Filés", "Filé de Porquinho", "500g", 21.00, "Sim", ""),
    ("Filés", "Filé de Salmão", "kg", 99.99, "Sim", ""),
    ("Filés", "Filé de Tilápia", "500g", 28.00, "Sim", ""),
    ("Filés", "Filé de Tilápia", "1kg", 55.00, "Sim", ""),
    ("Filés", "Filé de Tucunaré", "500g", 25.00, "Sim", ""),
    ("Filés", "Filé de Trilha", "500g", 21.00, "Sim", ""),
    ("Filés", "Filé de Uritinga (Pintado da Amazônia)", "1kg", 45.00, "Sim", ""),
    ("Filés", "Sardinha Espalmada", "500g", 15.00, "Sim", ""),
    ("Filés", "Traíra Espalmada Retalhada", "1kg", 36.00, "Sim", ""),
    ("Postas", "Costelinha de Pacu", "500g", 21.00, "Sim", ""),
    ("Postas", "Posta de Barbado", "500g", 16.50, "Sim", ""),
    ("Postas", "Posta de Cascudo", "500g", 14.00, "Sim", ""),
    ("Postas", "Posta de Cação", "Kg", 43.00, "Sim", ""),
    ("Postas", "Posta de Tilápia", "1kg", 35.00, "Sim", ""),
    ("Postas", "Posta de Pintado", "kg", 60.00, "Sim", ""),
    ("Peixes Inteiros", "Pacu Dessosado", "kg", 47.00, "Sim", ""),
    ("Peixes Inteiros", "Corvina (G)", "kg", 27.00, "Sim", ""),
    ("Peixes Inteiros", "Curimba", "kg", 18.00, "Sim", ""),
    ("Peixes Inteiros", "Lambari", "500g", 22.50, "Sim", ""),
    ("Peixes Inteiros", "Mandi", "kg", 24.00, "Sim", ""),
    ("Peixes Inteiros", "Manjubinha Eviscerada sem Cabeça", "500g", 20.00, "Sim", ""),
    ("Peixes Inteiros", "Pacu / Tambaqui com Escamas", "kg", 32.00, "Sim", ""),
    ("Peixes Inteiros", "Piapara", "kg", 25.00, "Sim", ""),
    ("Peixes Inteiros", "Piau 3 Pintas", "kg", 15.00, "Sim", ""),
    ("Peixes Inteiros", "Sardinha Inteira Eviscerada sem Cabeça", "500g", 12.50, "Sim", ""),
    ("Peixes Inteiros", "Tainha", "kg", 27.00, "Sim", ""),
    ("Peixes Inteiros", "Toquinho de Cascudo", "500g", 13.00, "Sim", ""),
    ("Peixes Inteiros", "Tilápia Eviscerada com Escamas", "kg", 27.00, "Sim", ""),
    ("Peixes Inteiros", "Tilápia Dessosada sem Escamas", "kg", 35.00, "Sim", ""),
    ("Peixes Inteiros", "Traíra", "kg", 20.00, "Sim", ""),
    ("Peixes Inteiros", "Traíra (G)", "kg", 27.00, "Sim", ""),
    ("Peixes Inteiros", "Tucunaré (G)", "kg", 30.00, "Sim", ""),
    ("Camarões", "Camarão (P)", "500g", 28.00, "Sim", ""),
    ("Camarões", "Camarão (M)", "500g", 39.50, "Sim", ""),
    ("Camarões", "Camarão (G)", "500g", 48.50, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Bacalhau", "500g", 23.90, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Bacalhau Gourmet", "500g", 30.00, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Camarão", "500g", 23.90, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Carne Seca com Mandioca", "500g", 23.90, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Costela", "500g", 23.90, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Salmão", "300g", 25.90, "Sim", ""),
    ("Prontos para Fritar", "Bolinho de Tilápia com Queijo", "500g", 23.90, "Sim", ""),
    ("Prontos para Fritar", "Camarão Empanado com Requeijão", "650g", 40.00, "Sim", ""),
    ("Prontos para Fritar", "Camarão Empanado", "400g", 43.00, "Sim", ""),
    ("Prontos para Fritar", "Chips de Tilápia Empanado", "500g", 28.50, "Sim", ""),
    ("Prontos para Fritar", "Espetinho de Tilápia com Bacon", "500g", 38.00, "Sim", ""),
    ("Prontos para Fritar", "Espetinho de Tilápia com Provolone", "500g", 38.00, "Sim", ""),
    ("Prontos para Fritar", "Filé de Porquinho Empanado", "500g", 30.00, "Sim", ""),
    ("Prontos para Fritar", "Iscas de Pintado Empanado", "500g", 33.00, "Sim", ""),
    ("Prontos para Fritar", "Lambari Desossado e Empanado", "330g", 28.00, "Sim", ""),
    ("Prontos para Fritar", "Moldado de Tilápia Empanado", "700g", 22.00, "Sim", ""),
    ("Prontos para Fritar", "Queijo Cura Empanado", "500g", 30.00, "Sim", ""),
    ("Prontos para Fritar", "Tiras de Merluza Empanada", "500g", 28.00, "Sim", ""),
    ("Prontos para Fritar", "Tirinhas de Pangasius Empanadas", "500g", 30.00, "Sim", ""),
    ("Prontos para Fritar", "Tirinhas de Tilápia Empanadas", "500g", 34.50, "Sim", ""),
    ("Prontos para Fritar", "Torresmo de Pacu Empanado", "500g", 36.00, "Sim", ""),
    ("Prontos para Fritar", "Iscas de Filé de Frango Empanado", "500g", 18.00, "Sim", ""),
    ("Diversos", "Anel de Lula", "500g", 36.00, "Sim", ""),
    ("Diversos", "Caranguejo", "6 unidades", 55.00, "Sim", ""),
    ("Diversos", "Farinha para Empanar", "400g", 14.00, "Sim", ""),
    ("Diversos", "Kani Kama", "250g", 13.00, "Sim", ""),
    ("Diversos", "Mandioca", "1kg", 10.00, "Sim", ""),
    ("Diversos", "Mexilhão Desconchado", "500g", 31.50, "Sim", ""),
    ("Diversos", "Sardinha Salgada", "400g", 18.00, "Sim", ""),
    ("Diversos", "Siri", "1Kg", 57.00, "Sim", ""),
]


def normalizar_texto(texto: str) -> str:
    """Remove acentos e padroniza para minúsculas para buscas naturais."""
    texto_sem_acento = unicodedata.normalize("NFKD", str(texto))
    texto_limpo = "".join(ch for ch in texto_sem_acento if not unicodedata.combining(ch))
    return " ".join(texto_limpo.lower().split())


def formatar_moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(".", ",")


def eh_vendido_por_kg(peso_str: str) -> bool:
    return str(peso_str).strip().lower() == "kg"


def gerar_chave_carrinho(produto: str, peso: str) -> str:
    """Evita duplicidade no carrinho por diferença de maiúsculas ou espaços."""
    return f"{normalizar_texto(produto)}||{normalizar_texto(peso)}"


def obter_foto_produto(row: pd.Series) -> str:
    img_planilha = str(row.get("Imagem", "")).strip()
    if img_planilha and img_planilha.lower() not in ["nan", "none", ""]:
        return img_planilha

    nome_produto = str(row.get("Produto", "")).strip()
    return FOTOS_ESPECIFICAS.get(nome_produto, "").strip()


@st.cache_data(ttl=60)
def carregar_produtos() -> pd.DataFrame:
    if LINK_PLANILHA_CSV.strip():
        df = pd.read_csv(LINK_PLANILHA_CSV)
    else:
        if not os.path.exists(ARQUIVO_EXCEL_LOCAL):
            df_novo = pd.DataFrame(
                DADOS_PADRAO,
                columns=["Categoria", "Produto", "Peso", "Preco", "Disponivel", "Imagem"],
            )
            with pd.ExcelWriter(ARQUIVO_EXCEL_LOCAL, engine="openpyxl") as writer:
                df_novo.to_excel(writer, sheet_name="Catalogo", index=False)
        df = pd.read_excel(ARQUIVO_EXCEL_LOCAL, sheet_name=0)

    df.columns = [str(c).strip() for c in df.columns]

    if "Disponivel" in df.columns:
        df = df[
            df["Disponivel"]
            .astype(str)
            .str.strip()
            .str.lower()
            .isin(["sim", "s", "yes", "1", "true"])
        ]

    # Tratamento seguro de preços: aceita "R$ 28,00" ou "28.00" e ignora linhas inválidas
    preco_limpo = (
        df["Preco"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace("r$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["Preco"] = pd.to_numeric(preco_limpo, errors="coerce")
    df = df.dropna(subset=["Preco"])
    df = df[df["Preco"] > 0]

    df["Categoria"] = df["Categoria"].astype(str).str.strip()
    df["Produto"] = df["Produto"].astype(str).str.strip()
    df["Peso"] = df["Peso"].astype(str).str.strip()

    return df.reset_index(drop=True)


df_produtos = carregar_produtos()

# ==========================================
# BANNER HERO MODERNO
# ==========================================
st.markdown(
    f"""<div class="hero-modern">
<div class="hero-pill">🌊 Catálogo Digital Oficial</div>
<h1>🐟 Peixaria <span>KL</span></h1>
<p>{SLOGAN}</p>
<div class="hero-features">
<span class="feature-chip">❄️ Frescor Garantido</span>
<span class="feature-chip">⚡ Pedido Rápido no WhatsApp</span>
<span class="feature-chip">📍 {ENDERECO_LOJA}</span>
</div>
</div>""",
    unsafe_allow_html=True,
)

# ==========================================
# BARRA DE AÇÕES RÁPIDAS + SELETOR DE TEMA
# ==========================================
c_top1, c_top2, c_top3, c_top4 = st.columns([2, 2, 1.1, 1])
with c_top1:
    st.link_button(
        f"💬 WhatsApp {TELEFONE_EXIBICAO}",
        f"https://wa.me/{WHATSAPP}?text={urllib.parse.quote('Olá, Peixaria KL! Vim pelo catálogo online e gostaria de tirar uma dúvida.')}",
        use_container_width=True,
    )
with c_top2:
    qtd_itens = sum(item["qtd"] for item in st.session_state.carrinho.values())
    subtotal_topo = sum(item["qtd"] * item["preco"] for item in st.session_state.carrinho.values())
    st.info(f"🛒 **Carrinho:** {qtd_itens} item(ns) • **Subtotal:** {formatar_moeda(subtotal_topo)}")
with c_top3:
    if st.button("🔄 Atualizar Preços", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with c_top4:
    label_tema = "☀️ Tema Claro" if st.session_state.modo_escuro else "🌙 Tema Escuro"
    if st.button(label_tema, use_container_width=True):
        st.session_state.modo_escuro = not st.session_state.modo_escuro
        st.rerun()

# ==========================================
# BARRA LATERAL: CARRINHO DE COMPRAS
# ==========================================
with st.sidebar:
    st.markdown("## 🛒 Seu Pedido")
    if not st.session_state.carrinho:
        st.info("Seu carrinho está vazio.\n\nEscolha os itens no catálogo ao lado e clique em **Adicionar**.")
    else:
        subtotal_produtos = 0.0
        tem_item_kg = False
        linhas_itens_zap = []

        for chave, item in list(st.session_state.carrinho.items()):
            subtotal = item["preco"] * item["qtd"]
            subtotal_produtos += subtotal
            por_kg = eh_vendido_por_kg(item["peso"])
            if por_kg:
                tem_item_kg = True

            etiqueta_peso = f"{item['peso']} • ⚖️ Estimativa / A pesar" if por_kg else item["peso"]
            st.markdown(f"**{item['produto']}** • `{etiqueta_peso}`")
            b1, b2, b3, b4 = st.columns([1, 1, 1, 2])
            if b1.button("➖", key=f"menos_{chave}"):
                if item["qtd"] > 1:
                    st.session_state.carrinho[chave]["qtd"] -= 1
                else:
                    del st.session_state.carrinho[chave]
                st.rerun()
            b2.markdown(f"**{item['qtd']}**")
            if b3.button("➕", key=f"mais_{chave}"):
                st.session_state.carrinho[chave]["qtd"] += 1
                st.rerun()
            b4.markdown(f"**{formatar_moeda(subtotal)}**")
            st.divider()

            if por_kg:
                linhas_itens_zap.append(
                    f"• {item['qtd']} peça(s)/estimativa de {item['produto']} ({formatar_moeda(item['preco'])}/kg — *pesar na balança*)"
                )
            else:
                linhas_itens_zap.append(
                    f"• {item['qtd']}x {item['produto']} ({item['peso']}) — {formatar_moeda(subtotal)}"
                )

        if tem_item_kg:
            st.warning(
                "⚖️ **Itens por kg:** O valor exibido é uma estimativa inicial. O peso exato será conferido na balança e confirmado pelo WhatsApp."
            )

        tipo_entrega = st.radio("Modo de recebimento:", ["🚚 Entrega (Delivery)", "🏪 Retirar na Loja"])
        eh_entrega = "Entrega" in tipo_entrega
        valor_entrega = TAXA_ENTREGA if eh_entrega else 0.0
        total_final = subtotal_produtos + valor_entrega

        st.markdown("---")
        label_subtotal = "Subtotal estimado:" if tem_item_kg else "Subtotal dos itens:"
        label_total = "Total estimado:" if tem_item_kg else "Total:"
        st.markdown(f"**{label_subtotal}** {formatar_moeda(subtotal_produtos)}")
        if eh_entrega:
            st.markdown(f"**Taxa de entrega:** + {formatar_moeda(valor_entrega)}")
        else:
            st.markdown("**Taxa de entrega:** Grátis (Retirada)")
            st.info(f"📍 **Retirada em:**\n{ENDERECO_LOJA}")
        st.markdown(f"### {label_total} {formatar_moeda(total_final)}")
        st.markdown("---")

        nome_cliente = st.text_input("Seu nome:", placeholder="Ex: Maria Oliveira")
        endereco = ""
        endereco_valido = True

        if eh_entrega:
            endereco = st.text_input(
                "Endereço de entrega (obrigatório):",
                placeholder="Rua, número, bairro...",
            )
            endereco_valido = len(endereco.strip()) >= 5
            if not endereco_valido:
                st.warning("📍 Informe o endereço completo (rua, número e bairro) para habilitar o envio da entrega.")

        forma_pagamento = st.selectbox("Pagamento:", ["Pix", "Cartão (Crédito/Débito)", "Dinheiro"])
        troco_para = ""
        if forma_pagamento == "Dinheiro":
            troco_para = st.text_input("Troco para quanto? (opcional)", placeholder="Ex: R$ 100,00")

        observacoes = st.text_input(
            "Observações:",
            placeholder="Ex: Peixe de aprox. 2kg, limpo em postas...",
        )

        # Montagem da mensagem para o WhatsApp
        linhas_zap = [f"Olá, *{NOME_LOJA}*! Gostaria de fazer o seguinte pedido:\n"]
        linhas_zap.extend(linhas_itens_zap)
        if tem_item_kg:
            linhas_zap.append("\n⚖️ _Pedido contém item(ns) por kg: favor confirmar o peso e o valor exato na balança._")
        linhas_zap.append(f"\n*{label_subtotal}* {formatar_moeda(subtotal_produtos)}")
        if eh_entrega:
            linhas_zap.append(f"*Taxa de entrega:* {formatar_moeda(valor_entrega)}")
        linhas_zap.append(f"*{label_total} {formatar_moeda(total_final)}*")
        linhas_zap.append(f"\n*Recebimento:* {tipo_entrega}")
        if not eh_entrega:
            linhas_zap.append(f"*Local de retirada:* {ENDERECO_LOJA}")
        if nome_cliente.strip():
            linhas_zap.append(f"*Cliente:* {nome_cliente.strip()}")
        if endereco.strip():
            linhas_zap.append(f"*Endereço:* {endereco.strip()}")
        linhas_zap.append(f"*Pagamento:* {forma_pagamento}")
        if forma_pagamento == "Dinheiro" and troco_para.strip():
            linhas_zap.append(f"*Troco para:* {troco_para.strip()}")
        if observacoes.strip():
            linhas_zap.append(f"*Obs:* {observacoes.strip()}")

        texto_final = urllib.parse.quote("\n".join(linhas_zap))
        link_whatsapp = f"https://wa.me/{WHATSAPP}?text={texto_final}"

        if endereco_valido:
            st.link_button("📲 Enviar Pedido no WhatsApp", link_whatsapp, type="primary", use_container_width=True)
        else:
            st.button("📲 Preencha o endereço para enviar", disabled=True, use_container_width=True)

        if st.button("🗑️ Limpar Carrinho", use_container_width=True):
            st.session_state.carrinho = {}
            st.rerun()

# ==========================================
# NAVEGAÇÃO POR CATEGORIAS E BUSCA
# ==========================================
st.markdown("### 🧭 Explore por Categoria")
contagem_cats = df_produtos["Categoria"].value_counts().to_dict()
categorias_disponiveis = ["Todas"] + list(df_produtos["Categoria"].unique())

if st.session_state.categoria_ativa not in categorias_disponiveis:
    st.session_state.categoria_ativa = "Todas"

cols_cat = st.columns(len(categorias_disponiveis))
for i, cat in enumerate(categorias_disponiveis):
    icone = ICONES_CATEGORIA.get(cat, "🌊")
    qtd_cat = len(df_produtos) if cat == "Todas" else contagem_cats.get(cat, 0)
    tipo_btn = "primary" if st.session_state.categoria_ativa == cat else "secondary"
    if cols_cat[i].button(f"{icone} {cat} ({qtd_cat})", key=f"cat_{cat}", type=tipo_btn, use_container_width=True):
        st.session_state.categoria_ativa = cat
        st.rerun()

busca = st.text_input(
    "🔍 O que você procura hoje?",
    placeholder="Pesquise por Salmao, Tilápia, Camarão, Bolinho, Posta...",
)

df_filtrado = df_produtos.copy()
if st.session_state.categoria_ativa != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Categoria"] == st.session_state.categoria_ativa]

if busca.strip():
    termo_busca = normalizar_texto(busca)
    mascara_busca = df_filtrado["Produto"].apply(lambda x: termo_busca in normalizar_texto(x)) | df_filtrado[
        "Categoria"
    ].apply(lambda x: termo_busca in normalizar_texto(x))
    df_filtrado = df_filtrado[mascara_busca]

df_filtrado = df_filtrado.reset_index(drop=True)
st.caption(f"Mostrando **{len(df_filtrado)}** produto(s) em **{st.session_state.categoria_ativa}**")

# ==========================================
# GRADE DE PRODUTOS (COM SUPORTE A FOTOS)
# ==========================================
if df_filtrado.empty:
    st.warning("Nenhum produto encontrado. Tente buscar por outro nome ou clique em 'Todas'.")
else:
    cols = st.columns(3)
    for idx, row in df_filtrado.iterrows():
        coluna = cols[idx % 3]
        with coluna:
            icone_cat = ICONES_CATEGORIA.get(row["Categoria"], "🐟")
            preco_fmt = formatar_moeda(row["Preco"])
            por_kg = eh_vendido_por_kg(row["Peso"])
            badge_peso = "⚖️ Por kg" if por_kg else f"📦 {row['Peso']}"
            label_preco = "Valor por kg (pesado na hora)" if por_kg else "Preço unitário"
            preco_exibicao = f"{preco_fmt}/kg" if por_kg else preco_fmt

            foto_url = obter_foto_produto(row)
            bloco_foto = ""
            if foto_url:
                bloco_foto = (
                    f'<div class="product-photo-box"><img src="{foto_url}" class="product-photo-img" '
                    f'alt="{row["Produto"]}" onerror="this.parentElement.style.display=\'none\'"></div>'
                )

            st.markdown(
                f"""<div class="product-card">
<div class="card-top-row">
<span class="cat-badge">{icone_cat} {row['Categoria']}</span>
<span class="weight-badge">{badge_peso}</span>
</div>
{bloco_foto}
<div class="prod-name">{row['Produto']}</div>
<div class="price-row">
<span class="price-label">{label_preco}</span>
<span class="price-number">{preco_exibicao}</span>
</div>
</div>""",
                unsafe_allow_html=True,
            )

            chave_item = gerar_chave_carrinho(row["Produto"], row["Peso"])
            if st.button("🛒 Adicionar", key=f"add_{chave_item}_{idx}", use_container_width=True):
                if chave_item in st.session_state.carrinho:
                    st.session_state.carrinho[chave_item]["qtd"] += 1
                else:
                    st.session_state.carrinho[chave_item] = {
                        "produto": row["Produto"],
                        "peso": row["Peso"],
                        "preco": float(row["Preco"]),
                        "qtd": 1,
                    }
                st.toast(f"✅ {row['Produto']} adicionado ao carrinho!")
                st.rerun()

# ==========================================
# RODAPÉ MODERNO
# ==========================================
st.markdown(
    f"""<div class="modern-footer">
<h3>🐟 {NOME_LOJA}</h3>
<p style="margin:4px 0; color:#e2e8f0;">📍 {ENDERECO_LOJA}</p>
<p style="margin:4px 0; color:#e2e8f0;">🕒 {HORARIO_ATENDIMENTO}</p>
<p style="margin:4px 0; color:#7dd3fc;">🚚 {CIDADE_REGIAO} • 📱 WhatsApp: <b>{TELEFONE_EXIBICAO}</b></p>
<p style="margin-top:14px; font-size:0.78rem; color:#94a3b8;">
- Valores sujeitos à disponibilidade de estoque. Itens por kg são pesados na balança e confirmados pelo WhatsApp. Taxa de entrega: {formatar_moeda(TAXA_ENTREGA)}.
</p>
</div>""",
    unsafe_allow_html=True,
)
