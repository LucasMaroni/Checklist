import streamlit as st
import os
from datetime import datetime
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import getpass

# Carregar variáveis do .env
load_dotenv()

# Configuração de página
st.set_page_config(page_title="Checklist de Caminhão", layout="centered")
st.title("🚚 Checklist de Caminhão")

# Estados da aplicação
if "etapa" not in st.session_state:
    st.session_state.etapa = 1
if "dados" not in st.session_state:
    st.session_state.dados = {}
if "imagens" not in st.session_state:
    st.session_state.imagens = []

# Função para enviar e-mail com anexos
def enviar_email(arquivo_word, arquivo_pdf):
    try:
        msg = EmailMessage()
        msg["Subject"] = f"Checklist - {st.session_state.dados['PLACA_CAMINHAO']}"
        msg["From"] = os.getenv("EMAIL_USER")
        msg["To"] = os.getenv("EMAIL_DESTINO")
        msg.set_content("Segue em anexo o checklist finalizado.")

        # Anexa Word
        msg.add_attachment(
            arquivo_word.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="Checklist_Preenchido.docx"
        )
        # Anexa PDF
        msg.add_attachment(
            arquivo_pdf.getvalue(),
            maintype="application",
            subtype="pdf",
            filename="Checklist_Final.pdf"
        )

        with smtplib.SMTP(os.getenv("EMAIL_HOST"), int(os.getenv("EMAIL_PORT"))) as smtp:
            smtp.starttls()
            smtp.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
            smtp.send_message(msg)

        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")
        return False

# Etapa 1
if st.session_state.etapa == 1:
    st.subheader("Etapa 1: Dados Básicos")
    st.session_state.dados['PLACA_CAMINHAO'] = st.text_input("Placa do Caminhão", max_chars=8, placeholder="ABC1234")
    st.session_state.dados['KM_ATUAL'] = st.text_input("KM Atual", placeholder="Ex: 120000")
    st.session_state.dados['MOTORISTA'] = st.text_input("Motorista", placeholder="Nome completo")
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

    tipo_carreta = st.radio("Carreta", ["2 EIXOS", "3 EIXOS"])
    st.session_state.dados["CARRETA_2"] = "X" if tipo_carreta == "2 EIXOS" else ""
    st.session_state.dados["CARRETA_3"] = "X" if tipo_carreta == "3 EIXOS" else ""

    if st.button("Avançar ➡️"):
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
        if st.button("Avançar ➡️"):
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

    if st.button("✅ Finalizar Checklist"):
        st.session_state.dados.update({k: "OK" if v else "NÃO OK" for k, v in checklist_itens.items()})
        st.session_state.dados['OBSERVACOES'] = observacao

        # Gerar Word em memória
        doc = Document("Checklist_Preenchivel.docx")
        for p in doc.paragraphs:
            for k, v in st.session_state.dados.items():
                if f"{{{{{k}}}}}" in p.text:
                    p.text = p.text.replace(f"{{{{{k}}}}}", str(v))
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for k, v in st.session_state.dados.items():
                            if f"{{{{{k}}}}}" in p.text:
                                p.text = p.text.replace(f"{{{{{k}}}}}", str(v))
        buffer_word = BytesIO()
        doc.save(buffer_word)
        buffer_word.seek(0)

        # Gerar PDF em memória
        buffer_pdf = BytesIO()
        c = canvas.Canvas(buffer_pdf, pagesize=A4)
        text = c.beginText(40, 800)
        text.setFont("Helvetica", 12)
        for chave, valor in st.session_state.dados.items():
            text.textLine(f"{chave}: {valor}")
        c.drawText(text)
        c.showPage()
        c.save()
        buffer_pdf.seek(0)

        # Enviar por e-mail
        if enviar_email(buffer_word, buffer_pdf):
            st.success("Checklist enviado por e-mail com sucesso!")

        # Disponibilizar para download
        st.download_button("📄 Baixar Word", buffer_word, file_name="Checklist_Preenchido.docx")
        st.download_button("📄 Baixar PDF", buffer_pdf, file_name="Checklist_Final.pdf")
