import streamlit as st
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import copy
from datetime import date
import os
import streamlit.components.v1 as components

# =========================
# LOGIN
# =========================
usuarios = ["Juan", "Bruno", "Josiel"]
senha_padrao = "BM123"

if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:
    st.title("🔐 Acesso Técnico")
    usuario = st.selectbox("Selecione seu nome", usuarios)
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if senha == senha_padrao:
            st.session_state.logado = True
            st.session_state.usuario = usuario
            st.success("✅ Acesso liberado!")
            st.rerun()
        else:
            st.error("❌ Senha incorreta")
    st.stop()

# =========================
# CONFIGURAÇÕES
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "modelo.docx")

MESES = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho",
    7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

# =========================
# FUNÇÕES DE SUBSTITUIÇÃO
# =========================
def full_text(p):
    return "".join(r.text for r in p.runs)

def substituir_texto_paragrafo(p, novo_texto):
    """Substitui texto preservando a formatação do primeiro run."""
    if not p.runs:
        return
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    primeiro = p.runs[0]._r
    rPr = primeiro.find(f'{{{ns}}}rPr')
    rPr_copia = copy.deepcopy(rPr) if rPr is not None else None

    for run in list(p.runs):
        run._r.getparent().remove(run._r)

    novo_run = p.add_run(novo_texto)
    if rPr_copia is not None:
        novo_run._r.insert(0, rPr_copia)

def inserir_bloco_texto(p_ref, linhas):
    """Substitui parágrafo {{TAG}} por múltiplas linhas."""
    parent = p_ref._element.getparent()
    idx = list(parent).index(p_ref._element)
    parent.remove(p_ref._element)

    for i, linha in enumerate(linhas):
        novo_p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = linha
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r.append(t)
        novo_p.append(r)
        parent.insert(idx + i, novo_p)

# =========================
# INTERFACE
# =========================
st.title("📄 Relatório Norte Energia")

tecnico = st.session_state.usuario
st.success(f"👷 Técnico logado: {tecnico}")

numero = st.text_input("Número do Relatório", "1")
assunto = st.text_input("Assunto", "MANUTENÇÃO RADAR CANAL DE FUGA")
data_manut = st.date_input("Data", value=date.today())
local = st.text_input("Localidade", "Canal de Fuga")

# =========================
# PARTE INFORMATIVA
# =========================
st.subheader("📝 Parte Informativa")

texto_informativo = st.text_area(
    "Editar se necessário:",
    f"""Manutenção Radar canal de fuga

A equipe de Meios Eletrônicos, sob a Superintendência de Segurança Corporativa, executou na data de {date.today().day} de {ME
