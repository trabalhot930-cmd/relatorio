import streamlit as st
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import date
import os
import urllib.parse
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
# TEMPLATE
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "modelo.docx")

# =========================
# MESES
# =========================
MESES = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril", 5:"Maio", 6:"Junho",
    7:"Julho", 8:"Agosto", 9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

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

A equipe de Meios Eletrônicos, sob a Superintendência de Segurança Corporativa, executou na data de {data_manut.day} de {MESES[data_manut.month].lower()} a manutenção do sistema radar da localidade canal de fuga.

Foram executadas as atividades de:
• Testes
• Religamento (equipamento estava congelado)
""",
    height=250
)

# =========================
# NOME DA IMAGEM
# =========================
st.subheader("🖼️ Identificação da Imagem")

nome_foto = st.text_input(
    "Descrição da imagem:",
    "Radar Canal de Fuga após manutenção"
)

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

    data_str = f"{data_manut.day} de {MESES[data_manut.month]} de {data_manut.year}"
    data_upper = data_str.upper()

    texto_ilustrativo = f"Manutenção em {local}"

    # =========================
    # SUBSTITUIÇÃO SEGURA
    # =========================
    for p in doc.paragraphs:

        if "RELATÓRIO DE SEGURANÇA" in p.text:
            p.text = f"RELATÓRIO DE SEGURANÇA Nr. {numero} / 2026"

        elif "DATA:" in p.text:
            p.text = f"DATA: {data_upper}"

        elif "ASSUNTO:" in p.text:
            p.text = f"ASSUNTO: {assunto.upper()}"

        elif "{{INFORMATIVA}}" in p.text:
            p.text = texto_informativo

        elif "{{ILUSTRATIVA}}" in p.text:
            p.text = texto_ilustrativo

        elif "{{CONCLUSIVA}}" in p.text:
            p.text = texto_conclusivo

    # =========================
    # FOTO + LEGENDA
    # =========================
    if foto_bytes:
        for i, p in enumerate(doc.paragraphs):
            if "PARTE ILUSTRATIVA" in p.text:
                try:
                    # Nome da imagem
                    doc.paragraphs[i+1].text = nome_foto
                    doc.paragraphs[i+1].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Imagem
                    run = doc.paragraphs[i+2].add_run()
                    run.add_picture(io.BytesIO(foto_bytes), width=Cm(16))
                    doc.paragraphs[i+2].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Legenda
                    doc.paragraphs[i+3].text = f"Figura 1 – {nome_foto}"
                    doc.paragraphs[i+3].alignment = WD_ALIGN_PARAGRAPH.CENTER

                except Exception as e:
                    st.error(f"Erro na imagem: {e}")

                break

    # =========================
    # DATA FINAL
    # =========================
    for p in doc.paragraphs:
        if "Vitória do Xingu" in p.text:
            p.text = f"Vitória do Xingu /PA, {data_str}"

    # =========================
    # EXPORTAR
    # =========================
    buffer = io.BytesIO()
    doc.save(buffer)

    st.success("✅ Relatório gerado com sucesso!")

    st.download_button(
        "📥 Baixar Relatório",
        buffer.getvalue(),
        f"Relatorio_{numero}.docx"
    )

    # =========================
    # WHATSAPP
    # =========================
    mensagem = f"""
Relatório Nr {numero}
Assunto: {assunto}
Local: {local}
"""

    link = "https://wa.me/?text=" + urllib.parse.quote(mensagem)

    st.markdown(f"### 📲 [Enviar via WhatsApp]({link})")
