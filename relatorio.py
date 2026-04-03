import streamlit as st
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
# CONFIG TEMPLATE
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "modelo.docx")

MESES = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho",
    7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

# =========================
# FUNÇÕES AUXILIARES
# =========================
def full_text(p):
    """Lê o texto completo de um parágrafo, mesmo com múltiplos runs."""
    return "".join(r.text for r in p.runs)

def substituir_paragrafo(p, novo_texto):
    """
    Substitui o texto de um parágrafo preservando a formatação (fonte, bold, etc.)
    do primeiro run. Funciona mesmo que o texto esteja dividido em vários runs.
    """
    if not p.runs:
        return

    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    primeiro_run = p.runs[0]._r
    rPr = primeiro_run.find(f'{{{ns}}}rPr')
    rPr_copia = copy.deepcopy(rPr) if rPr is not None else None

    # Remove todos os runs existentes
    for run in list(p.runs):
        run._r.getparent().remove(run._r)

    # Cria novo run com o texto novo
    novo_run = p.add_run(novo_texto)
    if rPr_copia is not None:
        novo_run._r.insert(0, rPr_copia)

def inserir_bloco_texto(p_ref, linhas):
    """
    Substitui um parágrafo marcador (ex: {{INFORMATIVA}}) por múltiplos
    parágrafos — um por linha do texto.
    """
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

numero    = st.text_input("Número do Relatório", "1")
assunto   = st.text_input("Assunto", "MANUTENÇÃO RADAR CANAL DE FUGA")
data_manut = st.date_input("Data", value=date.today())
local     = st.text_input("Localidade", "Canal de Fuga")

# =========================
# PARTE INFORMATIVA
# =========================
st.subheader("📝 Parte Informativa")

texto_informativo = st.text_area(
    "Editar se necessário:",
    f"""Manutenção Radar canal de fuga

A equipe de Meios Eletrônicos, sob a Superintendência de Segurança Corporativa, executou na data de {data_manut.day} de {MESES[data_manut.month].lower()} a manutenção do sistema radar da localidade canal de fuga.

Foram executadas as atividades de:
• Testes
• Religamento (equipamento estava congelado)
""",
    height=250
)

# =========================
# IMAGEM
# =========================
st.subheader("🖼️ Identificação da Imagem")
nome_foto = st.text_input("Descrição da imagem:", "Radar Canal de Fuga após manutenção")

# =========================
# PARTE CONCLUSIVA
# =========================
st.subheader("📌 Parte Conclusiva")
texto_conclusivo = st.text_area(
    "Editar conclusão:",
    "Após as manutenções os equipamentos foram recolocados em operação.",
    height=150
)

# =========================
# GPS
# =========================
st.subheader("📍 Localização automática")

gps_html = """
<script>
navigator.geolocation.getCurrentPosition(
    (position) => {
        const lat = position.coords.latitude;
        const lon = position.coords.longitude;
        const coords = lat + "," + lon;
        window.parent.postMessage({type: "streamlit:setComponentValue", value: coords}, "*");
    }
);
</script>
"""
coords = components.html(gps_html, height=0)
if coords:
    st.success(f"📍 GPS: {coords}")

# =========================
# FOTO
# =========================
st.subheader("📸 Foto da Atividade")
opcao_foto = st.radio("Escolha:", ["Tirar foto", "Enviar da galeria"])

foto_bytes = None
if opcao_foto == "Tirar foto":
    foto = st.camera_input("Abrir câmera")
    if foto:
        foto_bytes = foto.getvalue()
else:
    arquivo = st.file_uploader("Selecionar imagem", type=["jpg", "jpeg", "png"])
    if arquivo:
        foto_bytes = arquivo.read()

# =========================
# GERAR RELATÓRIO
# =========================
if st.button("🚀 Gerar Relatório"):

    if not os.path.exists(TEMPLATE_PATH):
        st.error("❌ modelo.docx não encontrado!")
        st.stop()

    doc = Document(TEMPLATE_PATH)

    data_str   = f"{data_manut.day} de {MESES[data_manut.month]} de {data_manut.year}"
    data_upper = data_str.upper()

    # -----------------------------------------------
    # 1) SUBSTITUIÇÕES NA TABELA DO CABEÇALHO
    #    (Nr. / DATA / ASSUNTO ficam numa tabela)
    # -----------------------------------------------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    txt = full_text(p)
                    if "RELATÓRIO DE SEGURANÇA" in txt:
                        substituir_paragrafo(p, f"RELATÓRIO DE SEGURANÇA Nr. {numero} / 2026")
                    elif "DATA:" in txt:
                        substituir_paragrafo(p, f"DATA: {data_upper}")
                    elif "ASSUNTO:" in txt:
                        substituir_paragrafo(p, f"ASSUNTO: {assunto.upper()}")

    # -----------------------------------------------
    # 2) SUBSTITUIÇÕES NOS PARÁGRAFOS DO CORPO
    # -----------------------------------------------
    # Fazemos snapshot da lista para não quebrar o loop
    # quando inserir_bloco_texto alterar o XML
    paragrafos = list(doc.paragraphs)

    for p in paragrafos:
        txt = full_text(p)

        if "{{INFORMATIVA}}" in txt:
            inserir_bloco_texto(p, texto_informativo.split('\n'))

        elif "{{ILUSTRATIVA}}" in txt:
            substituir_paragrafo(p, f"Manutenção em {local}")

        elif "{{CONCLUSIVA}}" in txt:
            inserir_bloco_texto(p, texto_conclusivo.split('\n'))

        elif "Vitória do Xingu" in txt:
            substituir_paragrafo(p, f"Vitória do Xingu /PA, {data_str}")

    # -----------------------------------------------
    # 3) FOTO + LEGENDA
    # -----------------------------------------------
    if foto_bytes:
        for i, p in enumerate(doc.paragraphs):
            if "PARTE ILUSTRATIVA" in full_text(p):
                try:
                    doc.paragraphs[i+1].clear()
                    doc.paragraphs[i+1].add_run(nome_foto)
                    doc.paragraphs[i+1].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    doc.paragraphs[i+2].clear()
                    run = doc.paragraphs[i+2].add_run()
                    run.add_picture(io.BytesIO(foto_bytes), width=Cm(16))
                    doc.paragraphs[i+2].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    doc.paragraphs[i+3].clear()
                    doc.paragraphs[i+3].add_run(f"Figura 1 – {nome_foto}")
                    doc.paragraphs[i+3].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    st.error(f"Erro ao inserir imagem: {e}")
                break

    # -----------------------------------------------
    # 4) EXPORTAR
    # -----------------------------------------------
    buffer = io.BytesIO()
    doc.save(buffer)

    st.success("✅ Relatório gerado com sucesso!")
    st.download_button(
        "📥 Baixar Relatório",
        buffer.getvalue(),
        f"Relatorio_{numero}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # WhatsApp
    mensagem = f"Relatório Nr {numero}\nAssunto: {assunto}\nLocal: {local}"
