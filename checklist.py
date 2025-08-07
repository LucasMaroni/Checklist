import streamlit as st
import os
from datetime import datetime
from docx import Document
from PyPDF2 import PdfMerger, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from io import BytesIO
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import getpass

# Carrega variáveis de ambiente do .env
load_dotenv()

CAMINHO_BASE = "checklists_salvos"  # Use uma pasta local
os.makedirs(CAMINHO_BASE, exist_ok=True)

st.set_page_config(page_title="Checklist de Caminhão", layout="centered")
st.title("Checklist de Caminhão")

if 'etapa' not in st.session_state:
    st.session_state.etapa = 1
if 'dados' not in st.session_state:
    st.session_state.dados = {}
if 'imagens' not in st.session_state:
    st.session_state.imagens = []

# Etapa 1
if st.session_state.etapa == 1:
    st.subheader("Etapa 1: Dados Básicos")
    st.session_state.dados['PLACA_CAMINHAO'] = st.text_input("Placa do Caminhão", max_chars=8)
    st.session_state.dados['KM_ATUAL'] = st.text_input("KM Atual")
    st.session_state.dados['MOTORISTA'] = st.text_input("Motorista")
    st.session_state.dados['PLACA_CARRETA1'] = st.text_input("Placa Carreta 1", max_chars=8)
    st.session_state.dados['PLACA_CARRETA2'] = st.text_input("Placa Carreta 2", max_chars=8)

    try:
        from ctypes import windll, create_unicode_buffer
        buffer = create_unicode_buffer(1024)
        windll.secur32.GetUserNameExW(3, buffer, (n := len(buffer)))
        nome_completo = buffer.value
    except:
        nome_completo = getpass.getuser()

    st.session_state.dados['VISTORIADOR'] = nome_completo

    tipo_veiculo = st.radio("Tipo de veículo", ["CAVALO", "RÍGIDO"])

    if tipo_veiculo == "CAVALO":
        subtipo = st.radio("Configuração do Cavalo", ["TOCO 4X2", "TRUCADO 6X2", "TRAÇADO 6X4"])
        st.session_state.dados.update({
            "CAVALO_TOCO": "X" if subtipo == "TOCO 4X2" else "",
            "CAVALO_TRUCADO": "X" if subtipo == "TRUCADO 6X2" else "",
            "CAVALO_TRACADO": "X" if subtipo == "TRAÇADO 6X4" else "",
            "RIGIDO_TOCO": "",
            "RIGIDO_TRUCADO": "",
            "RIGIDO_TRACADO": "",
        })
    else:
        subtipo = st.radio("Configuração do Rígido", ["TOCO 4X2", "TRUCADO 6X2", "TRAÇADO 6X4"])
        st.session_state.dados.update({
            "RIGIDO_TOCO": "X" if subtipo == "TOCO 4X2" else "",
            "RIGIDO_TRUCADO": "X" if subtipo == "TRUCADO 6X2" else "",
            "RIGIDO_TRACADO": "X" if subtipo == "TRAÇADO 6X4" else "",
            "CAVALO_TOCO": "",
            "CAVALO_TRUCADO": "",
            "CAVALO_TRACADO": "",
        })

    tipo_carreta = st.radio("CARRETA", ["2 EIXOS", "3 EIXOS"])
    st.session_state.dados["CARRETA_2"] = "X" if tipo_carreta == "2 EIXOS" else ""
    st.session_state.dados["CARRETA_3"] = "X" if tipo_carreta == "3 EIXOS" else ""

    if st.button("Avançar para Etapa 2"):
        if all([
            st.session_state.dados['PLACA_CAMINHAO'],
            st.session_state.dados['KM_ATUAL'],
            st.session_state.dados['MOTORISTA'],
            st.session_state.dados['PLACA_CARRETA1'],
            st.session_state.dados['PLACA_CARRETA2']
        ]):
            st.session_state.dados["DATA"] = datetime.now().strftime("%d/%m/%Y")
            st.session_state.dados["HORA"] = datetime.now().strftime("%H:%M")
            st.session_state.etapa = 2
        else:
            st.warning("Preencha todos os campos obrigatórios.")

# Etapa 2
elif st.session_state.etapa == 2:
    st.subheader("Etapa 2: Inserção de Fotos")
    imagens = st.file_uploader("Envie ao menos 4 fotos", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if imagens and len(imagens) >= 4:
        st.session_state.imagens = imagens
        if st.button("Avançar para Etapa 3"):
            st.session_state.etapa = 3
    else:
        st.warning("Envie no mínimo 4 imagens.")

# Etapa 3
elif st.session_state.etapa == 3:
    st.subheader("Etapa 3: Checklist Técnico")

    checklist_itens = {
        "VAZAMENTO_OLEO_MOTOR": st.checkbox("Vazamento de óleo motor"),
        "VAZAMENTO_AGUA_MOTOR": st.checkbox("Vazamento de água motor"),
        "OLEO_MOTOR_OK": st.checkbox("Nível de óleo de motor"),
        "ARREFECIMENTO_OK": st.checkbox("Nível do líquido de arrefecimento"),
        "OLEO_CAMBIO_OK": st.checkbox("Vazamento de óleo câmbio"),
        "OLEO_DIFERENCIAL_OK": st.checkbox("Vazamento de óleo diferencial"),
        "OLEO_CUBOS_OK": st.checkbox("Vazamento de óleo cubos"),
        "VAZAMENTO_AR_OK": st.checkbox("Vazamento de ar"),
        "PNEUS_OK": st.checkbox("Pneus avariados"),
        "PARABRISA_OK": st.checkbox("Para-brisa"),
        "ILUMINACAO_OK": st.checkbox("Iluminação"),
        "FAIXAS_REFLETIVAS_OK": st.checkbox("Faixas refletivas"),
        "FALHAS_PAINEL_OK": st.checkbox("Falhas no painel"),
        "FUNCIONAMENTO_TK_OK": st.checkbox("Funcionamento TK"),
        "TACOGRAFO_OK": st.checkbox("Funcionamento tacógrafo"),
        "FUNILARIA_OK": st.checkbox("Itens para funilaria"),
    }

    observacao = st.text_area("Observações")

    if st.button("Finalizar Checklist"):
        st.session_state.dados.update({k: "OK" if v else "NÃO OK" for k, v in checklist_itens.items()})
        st.session_state.dados['OBSERVACOES'] = observacao

        nome_pasta = f"{st.session_state.dados['PLACA_CAMINHAO']} - {datetime.now().strftime('%d.%m.%Y')}"
        caminho_pasta = os.path.join(CAMINHO_BASE, nome_pasta)
        os.makedirs(caminho_pasta, exist_ok=True)

        doc = Document("Checklist_Preenchivel.docx")

        for p in doc.paragraphs:
            inline = p.runs
            for k, v in st.session_state.dados.items():
                if f"{{{{{k}}}}}" in p.text:
                    full_text = p.text.replace(f"{{{{{k}}}}}", str(v))
                    for i in range(len(inline)):
                        inline[i].text = ""
                    inline[0].text = full_text

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        inline = p.runs
                        for k, v in st.session_state.dados.items():
                            if f"{{{{{k}}}}}" in p.text:
                                full_text = p.text.replace(f"{{{{{k}}}}}", str(v))
                                for i in range(len(inline)):
                                    inline[i].text = ""
                                inline[0].text = full_text

        caminho_docx_final = os.path.join(caminho_pasta, "Checklist_Preenchido.docx")
        doc.save(caminho_docx_final)

        for i, imagem in enumerate(st.session_state.imagens):
            caminho_img = os.path.join(caminho_pasta, f"foto_{i+1}.jpg")
            with open(caminho_img, "wb") as f:
                f.write(imagem.getbuffer())

        caminho_pdf = os.path.join(caminho_pasta, "Checklist_Final.pdf")
        c = canvas.Canvas(caminho_pdf, pagesize=A4)
        text = c.beginText(40, 800)
        text.setFont("Helvetica", 12)
        for chave, valor in st.session_state.dados.items():
            text.textLine(f"{chave}: {valor}")
        c.drawText(text)
        c.showPage()
        c.save()

        st.success("Checklist finalizado com sucesso! Arquivos salvos com sucesso.")
