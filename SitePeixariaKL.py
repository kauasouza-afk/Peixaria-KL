import os
import streamlit as st
import pandas as pd
import urllib.parse

# ==========================================
# CONFIGURAÇÕES DA PEIXARIA KL
# ==========================================
NOME_LOJA = "Peixaria KL"
WHATSAPP = "5516997179478"
TELEFONE_EXIBICAO = "(16) 99717-9478"
SLOGAN = "Peixes frescos, frutos do mar e porções selecionadas direto para sua mesa"
HORARIO_ATENDIMENTO = "Seg a Sáb: 08h às 18h • Dom: 08h às 12h"
CIDADE_REGIAO = "Matão e Região - SP"
TAXA_ENTREGA_AVISO = "Consulte taxa de entrega pelo WhatsApp"
LINK_PLANILHA_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQwKMzUp8ebSHgcxYUFm1sRIHqYIFlcpdhdiVfBqhUrKbX8zPCuCnbmWLaChwrSEtjA84oAAn9Tmcur/pub?gid=372842900&single=true&output=csv"

PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL_LOCAL = os.path.join(PASTA_ATUAL, "Catalogo_Peixaria_KL.xlsx")

st.set_page_config(
    page_title=f"{NOME_LOJA} | Catálogo Digital",
    page_icon="🐟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# DESIGN MODERNO (CSS CUSTOMIZADO)
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .stApp {
        background: linear-gradient(180deg, #eef6fc 0%, #f8fafc 40%, #f1f5f9 100%);
    }

    /* Esconde menu padrão superior do Streamlit para visual de site real */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}

    /* Banner Hero Moderno */
    .hero-modern {
        background: radial-gradient(circle at top right, #0ea5e9 0%, #0369a1 35%, #081b33 100%);
        padding: 38px 28px;
        border-radius: 24px;
        color: white;
        text-align: center;
        box-shadow: 0 20px 40px -15px rgba(8, 27, 51, 0.35);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
    }
    .hero-pill {
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
    }
    .hero-modern h1 {
        color: #ffffff !important;
        margin: 0;
        font-size: 2.9rem;
        font-weight: 800;
        letter-spacing: -1px;
        line-height: 1.1;
    }
    .hero-modern h1 span {
        background: linear-gradient(90deg, #38bdf8, #67e8f9);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .hero-modern p {
        margin: 12px auto 22px auto;
        max-width: 650px;
        font-size: 1.05rem;
        color: #e0f2fe !important;
        font-weight: 400;
    }
    .hero-features {
        display: flex;
        justify-content: center;
        flex-wrap: wrap;
        gap: 10px;
    }
    .feature-chip {
        background: rgba(8, 27, 51, 0.45);
        border: 1px solid rgba(125, 211, 252, 0.3);
        padding: 7px 15px;
        border-radius: 12px;
        font-size: 0.84rem;
        color: #f0f9ff !important;
        font-weight: 600;
    }

    /* Card de Produto Moderno */
    .product-card {
        background: #ffffff;
        border-radius: 18px;
        padding: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.04);
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        margin-bottom: 10px;
    }
    .product-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 14px 28px -8px rgba(2, 132, 199, 0.18);
        border-color: #7dd3fc;
    }
    .card-top-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
    }
    .cat-badge {
        background: #f0f9ff;
        color: #0284c7;
        font-size: 0.74rem;
        font-weight: 700;
        padding: 4px 11px;
        border-radius: 999px;
        border: 1px solid #bae6fd;
    }
    .weight-badge {
        background: #f8fafc;
        color: #475569;
        font-size: 0.76rem;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }

    /* Silhueta ilustrada em formato de peixe */
    .fish-badge-wrapper {
        display: flex;
        justify-content: center;
        align-items: center;
        margin: 10px 0 12px 0;
        filter: drop-shadow(0 4px 8px rgba(2, 132, 199, 0.2));
    }
    .fish-shape-art {
        width: 175px;
        height: 76px;
        display: flex;
        align-items: center;
        justify-content: center;
        padding-left: 16px;
        color: #ffffff;
        font-weight: 700;
        font-size: 0.8rem;
        letter-spacing: 0.4px;
        clip-path: polygon(
            0% 18%, 18% 42%, 35% 15%, 65% 8%, 88% 25%, 100% 50%, 
            88% 75%, 65% 92%, 35% 85%, 18% 58%, 0% 82%, 8% 50%
        );
        transition: transform 0.25s ease;
    }
    .product-card:hover .fish-shape-art {
        transform: scale(1.05);
    }

    .prod-name {
        color: #0f172a;
        font-size: 1.08rem;
        font-weight: 700;
        line-height: 1.35;
        min-height: 46px;
        margin: 6px 0 12px 0;
        text-align: center;
    }
    .price-row {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        padding-top: 12px;
        border-top: 1px solid #f1f5f9;
    }
    .price-label {
        font-size: 0.8rem;
        color: #64748b;
        font-weight: 500;
    }
    .price-number {
        font-size: 1.48rem;
        font-weight: 800;
        color: #0369a1;
        letter-spacing: -0.5px;
    }

    /* Botões Streamlit Estilizados */
    div.stButton > button {
        border-radius: 12px;
        font-weight: 600;
        transition: all 0.2s ease;
    }
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        border: none;
        color: white;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.25);
    }

    /* Rodapé Moderno */
    .modern-footer {
        background: #081b33;
        color: #cbd5e1;
        padding: 34px 24px;
        border-radius: 22px;
        text-align: center;
        margin-top: 48px;
        box-shadow: 0 -4px 20px rgba(8, 27, 51, 0.08);
    }
    .modern-footer h3 {
        color: #ffffff !important;
        font-weight: 800;
        margin-bottom: 8px;
    }

    /* Força cor escura nos títulos e textos gerais para não sumirem no modo escuro */
    .stApp h1, .stApp h2, .stApp h3, .stApp label, .stApp .stMarkdown p, .stApp [data-testid="stCaptionContainer"] {
        color: #0f172a !important;
    }
    /* Mantém o texto branco dentro do banner principal, barra lateral e rodapé */
    .hero-modern h1, .hero-modern p, .modern-footer h3, .modern-footer p, section[data-testid="stSidebar"] * {
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

ICONES_CATEGORIA = {
    "Filés": "🐟",
    "Postas": "🥩",
    "Peixes Inteiros": "🐠",
    "Camarões": "🦐",
    "Prontos para Fritar": "🍤",
    "Diversos": "🦀",
}

DADOS_PADRAO = [
    ("Filés", "Filé de Bacalhau", "kg", 57.00, "Sim"),
    ("Filés", "Filé de Cavalinha", "500g", 15.00, "Sim"),
    ("Filés", "Filé de Corvina", "500g", 19.00, "Sim"),
    ("Filés", "Filé de Corvina", "1kg", 37.00, "Sim"),
    ("Filés", "Filé de Merluza", "500g", 19.50, "Sim"),
    ("Filés", "Filézinho de Pacu", "500g", 21.00, "Sim"),
    ("Filés", "Filé de Pangasius", "1kg", 40.00, "Sim"),
    ("Filés", "Filé de Pirarucu em Pedaços", "500g", 26.00, "Sim"),
    ("Filés", "Filé de Porquinho", "500g", 21.00, "Sim"),
    ("Filés", "Filé de Salmão", "kg", 89.99, "Sim"),
    ("Filés", "Filé de Tilápia", "500g", 27.00, "Sim"),
    ("Filés", "Filé de Tilápia", "1kg", 53.00, "Sim"),
    ("Filés", "Filé de Tucunaré", "500g", 25.00, "Sim"),
    ("Filés", "Filé de Trilha", "500g", 21.00, "Sim"),
    ("Filés", "Filé de Uritinga (Pintado da Amazônia)", "1Kg", 45.00, "Sim"),
    ("Filés", "Sardinha Espalmada", "500g", 15.00, "Sim"),
    ("Filés", "Traíra Espalmada Retalhada", "kg", 36.00, "Sim"),
    ("Postas", "Costelinha de Pacu", "500g", 21.00, "Sim"),
    ("Postas", "Posta de Barbado", "500g", 16.50, "Sim"),
    ("Postas", "Posta de Cascudo", "500g", 14.00, "Sim"),
    ("Postas", "Posta de Cação", "Kg", 43.00, "Sim"),
    ("Postas", "Posta de Tilápia", "1kg", 35.00, "Sim"),
    ("Postas", "Posta de Pintado", "kg", 60.00, "Sim"),
    ("Peixes Inteiros", "Cavalinha Inteira", "kg", 20.00, "Sim"),
    ("Peixes Inteiros", "Corvina (G)", "kg", 27.00, "Sim"),
    ("Peixes Inteiros", "Curimba", "kg", 18.00, "Sim"),
    ("Peixes Inteiros", "Lambari", "500g", 21.00, "Sim"),
    ("Peixes Inteiros", "Mandi", "kg", 24.00, "Sim"),
    ("Peixes Inteiros", "Manjubinha Eviscerada sem Cabeça", "500g", 18.00, "Sim"),
    ("Peixes Inteiros", "Pacu / Tambaqui com Escamas", "kg", 30.00, "Sim"),
    ("Peixes Inteiros", "Piapara", "kg", 25.00, "Sim"),
    ("Peixes Inteiros", "Piau 3 Pintas", "kg", 15.00, "Sim"),
    ("Peixes Inteiros", "Sardinha Inteira Eviscerada sem Cabeça", "500g", 12.50, "Sim"),
    ("Peixes Inteiros", "Tainha", "kg", 27.00, "Sim"),
    ("Peixes Inteiros", "Toquinho de Cascudo", "500g", 13.00, "Sim"),
    ("Peixes Inteiros", "Tilápia Eviscerada com Escamas", "kg", 27.00, "Sim"),
    ("Peixes Inteiros", "Tilápia Eviscerada sem Escamas", "kg", 30.00, "Sim"),
    ("Peixes Inteiros", "Traíra", "kg", 20.00, "Sim"),
    ("Peixes Inteiros", "Traíra (G)", "kg", 27.00, "Sim"),
    ("Peixes Inteiros", "Tucunaré (G)", "kg", 30.00, "Sim"),
    ("Camarões", "Camarão (P)", "500g", 28.00, "Sim"),
    ("Camarões", "Camarão (M)", "500g", 39.50, "Sim"),
    ("Camarões", "Camarão (G)", "500g", 48.50, "Sim"),
    ("Prontos para Fritar", "Bolinho de Bacalhau", "500g", 23.90, "Sim"),
    ("Prontos para Fritar", "Bolinho de Bacalhau Gourmet", "500g", 30.00, "Sim"),
    ("Prontos para Fritar", "Bolinho de Camarão", "500g", 23.90, "Sim"),
    ("Prontos para Fritar", "Bolinho de Carne Seca com Mandioca", "500g", 23.90, "Sim"),
    ("Prontos para Fritar", "Bolinho de Costela", "500g", 23.90, "Sim"),
    ("Prontos para Fritar", "Bolinho de Salmão", "300g", 25.00, "Sim"),
    ("Prontos para Fritar", "Bolinho de Tilápia com Queijo", "500g", 23.90, "Sim"),
    ("Prontos para Fritar", "Camarão Empanado com Requeijão", "650g", 40.00, "Sim"),
    ("Prontos para Fritar", "Camarão Empanado", "400g", 43.00, "Sim"),
    ("Prontos para Fritar", "Chips de Tilápia Empanado", "500g", 28.50, "Sim"),
    ("Prontos para Fritar", "Espetinho de Tilápia com Bacon", "500g", 38.00, "Sim"),
    ("Prontos para Fritar", "Espetinho de Tilápia com Provolone", "500g", 38.00, "Sim"),
    ("Prontos para Fritar", "Filé de Porquinho Empanado", "500g", 30.00, "Sim"),
    ("Prontos para Fritar", "Iscas de Pintado Empanado", "500g", 30.00, "Sim"),
    ("Prontos para Fritar", "Lambari Desossado e Empanado", "330g", 25.50, "Sim"),
    ("Prontos para Fritar", "Moldado de Tilápia Empanado", "700g", 22.00, "Sim"),
    ("Prontos para Fritar", "Queijo Cura Empanado", "500g", 30.00, "Sim"),
    ("Prontos para Fritar", "Tiras de Merluza Empanada", "500g", 28.00, "Sim"),
    ("Prontos para Fritar", "Tirinhas de Pangasius Empanadas", "500g", 27.00, "Sim"),
    ("Prontos para Fritar", "Tirinhas de Tilápia Empanadas", "500g", 33.00, "Sim"),
    ("Prontos para Fritar", "Torresmo de Pacu Empanado", "500g", 32.00, "Sim"),
    ("Diversos", "Anel de Lula", "500g", 36.00, "Sim"),
    ("Diversos", "Caranguejo", "6 unidades", 55.00, "Sim"),
    ("Diversos", "Farinha para Empanar", "400g", 14.00, "Sim"),
    ("Diversos", "Kani Kama", "250g", 13.00, "Sim"),
    ("Diversos", "Mandioca", "1kg", 10.00, "Sim"),
    ("Diversos", "Mexilhão Desconchado", "500g", 31.50, "Sim"),
    ("Diversos", "Sardinha Salgada", "400g", 18.00, "Sim"),
    ("Diversos", "Siri", "1Kg", 57.00, "Sim"),
]

@st.cache_data(ttl=60)
def carregar_produtos():
    if LINK_PLANILHA_CSV.strip():
        df = pd.read_csv(LINK_PLANILHA_CSV)
    else:
        if not os.path.exists(ARQUIVO_EXCEL_LOCAL):
            df_novo = pd.DataFrame(DADOS_PADRAO, columns=["Categoria", "Produto", "Peso", "Preco", "Disponivel"])
            with pd.ExcelWriter(ARQUIVO_EXCEL_LOCAL, engine="openpyxl") as writer:
                df_novo.to_excel(writer, sheet_name="Catalogo", index=False)
        df = pd.read_excel(ARQUIVO_EXCEL_LOCAL, sheet_name="Catalogo")

    df.columns = [c.strip() for c in df.columns]
    if "Disponivel" in df.columns:
        df = df[df["Disponivel"].astype(str).str.strip().str.lower().isin(["sim", "s", "yes", "1"])]

    df["Preco"] = (
        df["Preco"]
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
        .astype(float)
    )
    return df.reset_index(drop=True)

df_produtos = carregar_produtos()

if "carrinho" not in st.session_state:
    st.session_state.carrinho = {}

# ==========================================
# BANNER HERO MODERNO
# ==========================================
st.markdown(f"""
    <div class="hero-modern">
        <div class="hero-pill">🌊 Catálogo Digital Oficial</div>
        <h1>🐟 Peixaria <span>KL</span></h1>
        <p>{SLOGAN}</p>
        <div class="hero-features">
            <span class="feature-chip">❄️ Frescor Garantido</span>
            <span class="feature-chip">⚡ Pedido Rápido no WhatsApp</span>
            <span class="feature-chip">📍 {CIDADE_REGIAO}</span>
        </div>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# BARRA DE AÇÕES RÁPIDAS
# ==========================================
c_top1, c_top2, c_top3 = st.columns([2, 2, 1])
with c_top1:
    st.link_button(
        f"💬 Atendimento WhatsApp {TELEFONE_EXIBICAO}",
        f"https://wa.me/{WHATSAPP}?text={urllib.parse.quote('Olá, Peixaria KL! Vim pelo catálogo online e gostaria de tirar uma dúvida.')}",
        use_container_width=True
    )
with c_top2:
    qtd_itens = sum(item["qtd"] for item in st.session_state.carrinho.values())
    total_topo = sum(item["qtd"] * item["preco"] for item in st.session_state.carrinho.values())
    st.info(f"🛒 **Carrinho:** {qtd_itens} item(ns) • **R$ {total_topo:,.2f}**".replace(".", ","))
with c_top3:
    if st.button("🔄 Atualizar Preços", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ==========================================
# BARRA LATERAL: CARRINHO DE COMPRAS
# ==========================================
with st.sidebar:
    st.markdown("## 🛒 Seu Pedido")
    if not st.session_state.carrinho:
        st.info("Seu carrinho está vazio.\n\nEscolha os itens no catálogo ao lado e clique em **Adicionar**.")
    else:
        total = 0.0
        linhas_zap = [f"Olá, *{NOME_LOJA}*! Gostaria de fazer o seguinte pedido:\n"]

        for chave, item in list(st.session_state.carrinho.items()):
            subtotal = item["preco"] * item["qtd"]
            total += subtotal

            st.markdown(f"**{item['produto']}** • `{item['peso']}`")
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
            b4.markdown(f"**R$ {subtotal:,.2f}**".replace(".", ","))
            st.divider()

            linhas_zap.append(f"• {item['qtd']}x {item['produto']} ({item['peso']}) - R$ {subtotal:,.2f}".replace(".", ","))

        st.markdown(f"### Total: R$ {total:,.2f}".replace(".", ","))
        st.caption(f"ℹ️ {TAXA_ENTREGA_AVISO}")

        tipo_entrega = st.radio("Recebimento:", ["🚚 Entrega (Delivery)", "🏪 Retirar na Loja"])
        nome_cliente = st.text_input("Seu nome:", placeholder="Ex: Maria Oliveira")
        endereco = ""
        if "Entrega" in tipo_entrega:
            endereco = st.text_input("Endereço de entrega:", placeholder="Rua, número, bairro...")
        forma_pagamento = st.selectbox("Pagamento:", ["Pix", "Cartão (Crédito/Débito)", "Dinheiro"])
        observacoes = st.text_input("Observações:", placeholder="Ex: Limpo em postas, troco para R$ 100...")

        linhas_zap.append(f"\n*Total estimado: R$ {total:,.2f}*".replace(".", ","))
        linhas_zap.append(f"*Recebimento:* {tipo_entrega}")
        if nome_cliente.strip():
            linhas_zap.append(f"*Cliente:* {nome_cliente.strip()}")
        if endereco.strip():
            linhas_zap.append(f"*Endereço:* {endereco.strip()}")
        linhas_zap.append(f"*Pagamento:* {forma_pagamento}")
        if observacoes.strip():
            linhas_zap.append(f"*Obs:* {observacoes.strip()}")

        texto_final = urllib.parse.quote("\n".join(linhas_zap))
        link_whatsapp = f"https://wa.me/{WHATSAPP}?text={texto_final}"

        st.link_button("📲 Enviar Pedido no WhatsApp", link_whatsapp, type="primary", use_container_width=True)

        if st.button("🗑️ Limpar Carrinho", use_container_width=True):
            st.session_state.carrinho = {}
            st.rerun()

# ==========================================
# NAVEGAÇÃO POR CATEGORIAS E BUSCA
# ==========================================
st.markdown("### 🧭 Explore por Categoria")
categorias_disponiveis = ["Todas"] + list(df_produtos["Categoria"].unique())

if "categoria_ativa" not in st.session_state:
    st.session_state.categoria_ativa = "Todas"

cols_cat = st.columns(len(categorias_disponiveis))
for i, cat in enumerate(categorias_disponiveis):
    icone = ICONES_CATEGORIA.get(cat, "🌊")
    tipo_btn = "primary" if st.session_state.categoria_ativa == cat else "secondary"
    if cols_cat[i].button(f"{icone} {cat}", key=f"cat_{cat}", type=tipo_btn, use_container_width=True):
        st.session_state.categoria_ativa = cat
        st.rerun()

busca = st.text_input("🔍 O que você procura hoje?", placeholder="Pesquise por Salmão, Tilápia, Camarão, Bolinho, Posta...")

df_filtrado = df_produtos.copy()
if st.session_state.categoria_ativa != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Categoria"] == st.session_state.categoria_ativa]
if busca.strip():
    df_filtrado = df_filtrado[df_filtrado["Produto"].str.contains(busca, case=False, na=False)]

st.caption(f"Mostrando **{len(df_filtrado)}** produto(s) em **{st.session_state.categoria_ativa}**")

# ==========================================
# GRADE DE PRODUTOS
# ==========================================
if df_filtrado.empty:
    st.warning("Nenhum produto encontrado. Tente buscar por outro nome ou clique em 'Todas'.")
else:
    cols = st.columns(3)
    for idx, row in df_filtrado.iterrows():
        coluna = cols[idx % 3]
        with coluna:
            icone_cat = ICONES_CATEGORIA.get(row["Categoria"], "🐟")
            preco_fmt = f"R$ {row['Preco']:,.2f}".replace(".", ",")
            
            cat_lower = str(row.get("Categoria", "")).lower()
            prod_lower = str(row.get("Produto", "")).lower()
            if "salm" in prod_lower or "posta" in cat_lower:
                cor_peixe = "linear-gradient(135deg, #f97316, #ea580c)"
            elif "camar" in cat_lower or "camar" in prod_lower:
                cor_peixe = "linear-gradient(135deg, #ec4899, #db2777)"
            elif "pronto" in cat_lower:
                cor_peixe = "linear-gradient(135deg, #eab308, #ca8a04)"
            else:
                cor_peixe = "linear-gradient(135deg, #0284c7, #0369a1)"

            st.markdown(f"""
                <div class="product-card">
                    <div class="card-top-row">
                        <span class="cat-badge">{icone_cat} {row['Categoria']}</span>
                        <span class="weight-badge">📦 {row['Peso']}</span>
                    </div>
                    <div class="fish-badge-wrapper">
                        <div class="fish-shape-art" style="background: {cor_peixe};">
                            {icone_cat} {row['Peso']}
                        </div>
                    </div>
                    <div class="prod-name">{row['Produto']}</div>
                    <div class="price-row">
                        <span class="price-label">Preço unitário</span>
                        <span class="price-number">{preco_fmt}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            if st.button("🛒 Adicionar", key=f"add_{idx}", use_container_width=True):
                chave = f"{row['Produto']}_{row['Peso']}"
                if chave in st.session_state.carrinho:
                    st.session_state.carrinho[chave]["qtd"] += 1
                else:
                    st.session_state.carrinho[chave] = {
                        "produto": row["Produto"],
                        "peso": row["Peso"],
                        "preco": row["Preco"],
                        "qtd": 1,
                    }
                st.toast(f"✅ {row['Produto']} adicionado ao carrinho!")
                st.rerun()

# ==========================================
# RODAPÉ MODERNO
# ==========================================
st.markdown(f"""
    <div class="modern-footer">
        <h3>🐟 {NOME_LOJA}</h3>
        <p style="margin:4px 0; color:#e2e8f0;">{HORARIO_ATENDIMENTO}</p>
        <p style="margin:4px 0; color:#7dd3fc;">📍 {CIDADE_REGIAO} • 📱 WhatsApp: <b>{TELEFONE_EXIBICAO}</b></p>
        <p style="margin-top:14px; font-size:0.78rem; color:#94a3b8;">
            * Imagens e valores meramente ilustrativos, sujeitos à disponibilidade de estoque.
        </p>
    </div>
""", unsafe_allow_html=True)
