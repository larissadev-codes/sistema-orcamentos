# CHATBOT ORÇAMENTISTA
# Spraying Systems Co.

# 📊 FLUXO DE USO COM A PLANILHA
# 1. Realize a busca do produto no chatbot
# 2. Copie a linha exibida (Ctrl+C no texto)
# 3. Cole na coluna "Descrição" da planilha (Ctrl+V)
# 4. Preencha manualmente a coluna "Unid.": M → Tubos | PÇ → Conexões
# 5. O Excel calculará o custo total automaticamente

import re
import pandas as pd
import streamlit as st
import unicodedata
import urllib.parse

from datetime import datetime
from difflib import get_close_matches

data_atualizacao = datetime(2026, 5, 19)

# ─────────────────────────────────────────────
# LER ARQUIVO
# ─────────────────────────────────────────────

@st.cache_data
def carregar_dados():
    return pd.read_csv("Produtos.csv")

df = carregar_dados()


# ─────────────────────────────────────────────
# CORREÇÃO DE TEXTO
# ─────────────────────────────────────────────

def normalizar_texto(texto):
    texto = texto.lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = texto.encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"[^\w\s]", "", texto)
    return texto

def corrigir_palavras(busca, lista_palavras):
    palavras = busca.split()
    palavras_corrigidas = []
    for p in palavras:
        match = get_close_matches(p, lista_palavras, n=1, cutoff=0.7)
        palavras_corrigidas.append(match[0] if match else p)
    return " ".join(palavras_corrigidas)

def corrigir_tokens_complexos(busca):
    busca = re.sub(r'sxh\s*(\d+)', r'sch \1', busca)
    busca = re.sub(r'sc\s*(\d+)',  r'sch \1', busca)
    busca = re.sub(r'sh\s*(\d+)',  r'sch \1', busca)
    busca = re.sub(r'dm\s*(\d+)',  r'dn \1',  busca)
    busca = re.sub(r'dn\s*(\d+)',  r'dn \1',  busca)
    return busca

# ─────────────────────────────────────────────
# VERIFICAÇÃO DA BASE
# ─────────────────────────────────────────────

def verificar_base():
    hoje = datetime.today()
    dias = (hoje - data_atualizacao).days
    return dias > 180, dias

# ─────────────────────────────────────────────
# EMAIL AUTOMÁTICO
# ─────────────────────────────────────────────

def gerar_link_email(descricao):
    """
    Gera link mailto com destinatário correto por tipo de produto.
    Tubos → anselmo@caporal.com.br
    Conexões → dorattioto@supremy.com.br
    """
    descricao_lower = descricao.lower()

    if "tubo" in descricao_lower:
        destinatario = "anselmo@caporal.com.br"
    else:
        destinatario = "dorattioto@supremy.com.br"

    assunto = f"Solicitação de cotação - {descricao}"

    corpo = (
        f"Olá,\n\n"
        f"Por gentileza, orçar conforme descrição abaixo:\n"
        f"{descricao} - QUANTIDADE:\n\n"
        f"Agradeço desde já e aguardo o retorno.\n"
    )

    # ✅ CC com vírgula (padrão mailto) e & antes de subject
    link = (
        f"mailto:{destinatario}"
        f"?cc=l.luca@spray.com.br,wagner@spray.com.br,a.theodoro@spray.com.br,l.barros@spray.com.br"
        f"&subject={urllib.parse.quote(assunto)}"
        f"&body={urllib.parse.quote(corpo)}"
    )

    return link

# ─────────────────────────────────────────────
# FORMATAR LINHA PARA EXCEL
# ─────────────────────────────────────────────

def formatar_linha_excel(codigo, descricao, custo):
    """
    Ordem: Descrição | Código | Classificação | Unid. | Qtd | Custo Unit.
    Unidade deixada vazia — preencher manualmente (M ou PÇ).
    """
    descricao_lower = descricao.lower()
    unidade = "M" if "tubo" in descricao_lower else "PÇ"
    custo_fmt = f"{custo:.2f}".replace(".", ",")
    return f"{descricao}\t{codigo}\tMecânico\t{unidade}\t1\t{custo_fmt}"

# ─────────────────────────────────────────────
# BUSCA DE PRODUTOS  (sem st.button aqui dentro!)
# ─────────────────────────────────────────────

def buscar_produtos(busca):

    if not busca or busca.strip() == "":
        return "⚠️ Digite uma busca válida."

    busca_original = busca
    busca_normalizada = normalizar_texto(busca)

    # PERGUNTA DE ATUALIZAÇÃO
    if "atualizacao" in busca_normalizada:
        desatualizado, dias = verificar_base()
        data_formatada = data_atualizacao.strftime("%d/%m/%Y")
        if desatualizado:
            return (
                f"📅 Última atualização: {data_formatada}\n\n"
                f"<span style='color:red;font-weight:bold;'>⚠️ Base desatualizada há {dias} dias.</span>"
            )
        else:
            return (
                f"📅 Última atualização: {data_formatada}\n\n"
                f"<span style='color:green;font-weight:bold;'>✅ Base dentro do prazo.</span>"
            )

    # BUSCA DIRETA POR CÓDIGO
    match = re.search(r"\b[A-Z]{3,}\d{5,}\b", busca_original.upper())
    if match and " " not in busca_original:
        codigo = match.group()
        resultado = df[df['CODIGO'].astype(str) == codigo]
        if resultado.empty:
            return {"status": "nao_encontrado", "busca": busca_original}
        return {"status": "encontrado", "df": resultado}

    # PROCESSAMENTO DO TEXTO
    busca = busca.lower().replace('"', ' ')
    palavras_ignorar = ["aisi"]
    for palavra in palavras_ignorar:
        busca = busca.replace(palavra, "")

    busca = re.sub(r"([a-z]+)(dn\d+)",  r"\1 \2", busca)
    busca = re.sub(r"(dn\d+)(sch\d+)",  r"\1 \2", busca)
    busca = re.sub(r"dn\s*(\d+)",        r"dn \1",  busca)
    busca = re.sub(r"sch\s*(\d+)",       r"sch \1", busca)
    busca = corrigir_tokens_complexos(busca)
    
    PALAVRAS_BUSCA = [
        "tubo",
        "flange",
        "sorf",
        "solto",
        "cap",
        "curva",
        "luva",
        "meia",
        "304",
        "316",
        "dn",
        "sch"
        ]

    busca = corrigir_palavras(busca, PALAVRAS_BUSCA)

    busca = re.sub(r"\s+", " ", busca).strip()

    palavras = busca.split()

    # FILTROS
    filtro = pd.Series([True] * len(df), index=df.index)
    i = 0
    while i < len(palavras):
        p = palavras[i]

        if p == "dn" and i + 1 < len(palavras):
            dn_valor = palavras[i + 1].strip()
            filtro = filtro & (
                df["DN"].astype(str).str.replace('"', '', regex=False).str.strip() == dn_valor
            )
            i += 2
            continue

        if p == "sch" and i + 1 < len(palavras):
            sch_valor = palavras[i + 1].strip()
            filtro = filtro & (
                df["SCH"].astype(str).str.strip().str.extract(r"(\d+)")[0] == sch_valor
            )
            i += 2
            continue

        if p in ["304", "316"]:
            filtro = filtro & df["MATERIAL"].str.lower().str.contains(p, na=False)
            i += 1
            continue

        if p not in ["dn", "sch"] and not p.isdigit():
            filtro = filtro & df["DESCRICAO"].str.lower().str.contains(p, na=False)
            i += 1
            continue

        i += 1

    resultado = df[filtro]

    if resultado.empty:
        return {"status": "nao_encontrado", "busca": busca_original}

    return {"status": "encontrado", "df": resultado}

# ─────────────────────────────────────────────
# INTERFACE — LAYOUT
# ─────────────────────────────────────────────

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("spraying_systems_LOGO.jpg", width=300)

st.markdown(
    "<h1 style='text-align:center;color:#1f77b4;'>🏢 Sistema de Orçamentos</h1>",
    unsafe_allow_html=True
)
st.markdown(
    "<h3 style='text-align:center;color:gray;'>📊 Consulta Inteligente de Materiais</h3>",
    unsafe_allow_html=True
)
st.divider()

st.markdown(
    "<p style='color:gray; font-size:13px; margin-bottom:6px;'>💡 Exemplos de busca:</p>",
    unsafe_allow_html=True
)

exemplos = [
    "tubo dn 1-1/2 sch 40 aisi 304",
    "tubo dn 2 sch 40 aisi 304",
    "meia luva dn 1/4 aisi 304",
    "flange sorf dn 2 aisi 304",
    "flange solto dn 2 aisi 304"
]

cols = st.columns(len(exemplos))

for i, exemplo in enumerate(exemplos):
    with cols[i]:
        if st.button(exemplo, key=f"chip_{i}"):
            st.session_state.ultima_busca = exemplo
            st.session_state.ultima_resposta = buscar_produtos(exemplo)
            st.rerun()

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────

if "ultima_busca" not in st.session_state:
    st.session_state.ultima_busca = None
if "ultima_resposta" not in st.session_state:
    st.session_state.ultima_resposta = None

# ─────────────────────────────────────────────
# BUSCA
# ─────────────────────────────────────────────

prompt = st.chat_input("Digite sua busca...")

if prompt:
    st.session_state.ultima_busca = prompt
    st.session_state.ultima_resposta = buscar_produtos(prompt)

# ─────────────────────────────────────────────
# EXIBIR PERGUNTA DO USUÁRIO
# ─────────────────────────────────────────────

if st.session_state.ultima_busca:
    with st.chat_message("user"):
        st.markdown(st.session_state.ultima_busca)

# ─────────────────────────────────────────────
# EXIBIR RESPOSTA
# ─────────────────────────────────────────────

resposta = st.session_state.ultima_resposta

if resposta:
    with st.chat_message("assistant"):

        # ── TEXTO SIMPLES (atualização da base) ──
        if isinstance(resposta, str):
            st.markdown(resposta, unsafe_allow_html=True)

        # ── DICT ──
        elif isinstance(resposta, dict):

            # NÃO ENCONTRADO
            if resposta["status"] == "nao_encontrado":
                st.error("❌ Produto não encontrado na base.")
                st.markdown("Quer solicitar uma cotação para esse item?")

                link_email = gerar_link_email(resposta.get("busca", "Item não identificado"))

                st.markdown(f"""
                    <a href="{link_email}" target="_blank">
                        <button style="
                            background-color:#1f77b4;
                            color:white;
                            border:none;
                            padding:10px 18px;
                            border-radius:6px;
                            font-size:15px;
                            cursor:pointer;
                            margin-top:8px;
                        ">📧 Solicitar cotação</button>
                    </a>
                """, unsafe_allow_html=True)

            # ENCONTRADO
            elif resposta["status"] == "encontrado":

                # Alerta de base
                desatualizado, dias = verificar_base()
                if desatualizado:
                    st.markdown(
                        f"<div style='color:red;font-weight:bold;'>⚠️ Base desatualizada há {dias} dias.</div>",
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        "<div style='color:green;font-weight:bold;'>✅ Base dentro do prazo.</div>",
                        unsafe_allow_html=True
                    )

                st.markdown(
                    "<div style='color:#0b5394;font-weight:bold;font-size:18px;margin-top:8px;'>✅ Resultado encontrado:</div>",
                    unsafe_allow_html=True
                )

                resultado_df = resposta["df"]

                for _, linha in resultado_df.iterrows():
 
                    custo_raw = linha["CUSTO"]

                    try:
                        custo = float(str(custo_raw).replace(",", ".").strip())
                    except:
                        custo = None

                    tem_custo = custo is not None and custo > 0

                    st.markdown(f"🔹 **Código:** {linha['CODIGO']}")
                    st.markdown(f"📄 **Descrição:** {linha['DESCRICAO']}")

                    if tem_custo:
                        # ── COM CUSTO: mostra valor e linha para copiar ──
                        st.markdown(f"💰 **Custo:** R$ {custo:.2f}".replace(".", ","))

                        linha_excel = formatar_linha_excel(
                            linha["CODIGO"],
                            linha["DESCRICAO"],
                            custo
                        )

                        st.text_area(
                            "📋 Copiar linha para Excel (selecione e Ctrl+C):",
                            value=linha_excel,
                            key=f"copy_{linha['CODIGO']}",
                            height=80
                        )

                    else:
                        # ── SEM CUSTO: mostra aviso e botão de e-mail ──
                        st.markdown("💰 **Custo:** ⚠️ Não disponível")

                        link_email = gerar_link_email(
                            f"{linha['DESCRICAO']} - {linha['CODIGO']}"
                        )

                        st.markdown(f"""
                            <a href="{link_email}" target="_blank">
                                <button style="
                                    background-color:#1f77b4;
                                    color:white;
                                    border:none;
                                    padding:10px 18px;
                                    border-radius:6px;
                                    font-size:15px;
                                    cursor:pointer;
                                    margin-top:6px;
                                ">📧 Solicitar cotação</button>
                            </a>
                        """, unsafe_allow_html=True)

                    st.divider()
