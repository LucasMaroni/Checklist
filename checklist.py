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

# Carregar variáveis de ambiente
load_dotenv()


# Configuração da página
st.set_page_config(page_title="Checklist de Caminhão", layout="centered")
col_esq, col_dir = st.columns([4, 1])
with col_dir:
    st.image("logo.jpg", width=100)  # Aqui define largura fixa de 150 pixels
st.title("🚚 CheckList Manutenção")


# Estados
if "etapa" not in st.session_state:
    st.session_state.etapa = 1
if "dados" not in st.session_state:
    st.session_state.dados = {}
if "imagens" not in st.session_state:
    st.session_state.imagens = []
if "fotos_nao_ok" not in st.session_state:
    st.session_state.fotos_nao_ok = {}

# Função envio de e-mail
def enviar_email(arquivo_word, arquivo_pdf, fotos_extra):
    try:
        msg = EmailMessage()
        msg["Subject"] = f"Checklist - {st.session_state.dados['PLACA_CAMINHAO']}"
        msg["From"] = os.getenv("EMAIL_USER")
        msg["To"] = os.getenv("EMAIL_DESTINO")
        msg.set_content("Segue em anexo o checklist finalizado e imagens.")

        msg.add_attachment(
            arquivo_word.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="Checklist_Preenchido.docx"
        )
        msg.add_attachment(
            arquivo_pdf.getvalue(),
            maintype="application",
            subtype="pdf",
            filename="Checklist_Final.pdf"
        )
        for nome_item, arquivo in fotos_extra.items():
            if arquivo:
                msg.add_attachment(
                    arquivo.getvalue(),
                    maintype="image",
                    subtype="jpeg",
                    filename=f"foto_{nome_item}.jpg"
                )

        with smtplib.SMTP(os.getenv("EMAIL_HOST"), int(os.getenv("EMAIL_PORT"))) as smtp:
            smtp.starttls()
            smtp.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
            smtp.send_message(msg)

        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")
        return False

# -----------------
# ETAPA 1
# -----------------
if st.session_state.etapa == 1:
    st.subheader("Dados do Veículo e Condutor")
    st.session_state.dados['PLACA_CAMINHAO'] = st.text_input("Placa do Caminhão", max_chars=8)
    st.session_state.dados['KM_ATUAL'] = st.text_input("KM Atual")
    st.session_state.dados['MOTORISTA'] = st.text_input("Motorista")

    try:
        from ctypes import windll, create_unicode_buffer
        buffer = create_unicode_buffer(1024)
        windll.secur32.GetUserNameExW(3, buffer, (n := len(buffer)))
        nome_completo = buffer.value
    except:
        nome_completo = getpass.getuser()
    st.session_state.dados['VISTORIADOR'] = nome_completo

    tipo_veiculo = st.radio("Tipo de veículo", ["CAVALO", "RÍGIDO"])
    st.session_state.dados['TIPO_VEICULO'] = tipo_veiculo

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

        bitrem = st.toggle("Veículo é BITREM?")
        st.session_state.dados['BITREM'] = "SIM" if bitrem else "NÃO"
        st.session_state.dados['PLACA_CARRETA1'] = st.text_input("Placa Carreta 1", max_chars=8)

        if bitrem:
            st.session_state.dados['PLACA_CARRETA2'] = st.text_input("Placa Carreta 2", max_chars=8)
        else:
            st.session_state.dados['PLACA_CARRETA2'] = ""

        tipo_carreta = st.radio("Tipo de Carreta", ["2 EIXOS", "3 EIXOS"])
        st.session_state.dados["CARRETA_2"] = "X" if tipo_carreta == "2 EIXOS" else ""
        st.session_state.dados["CARRETA_3"] = "X" if tipo_carreta == "3 EIXOS" else ""

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
        st.session_state.dados['PLACA_CARRETA1'] = ""
        st.session_state.dados['PLACA_CARRETA2'] = ""
        st.session_state.dados['BITREM'] = "NÃO"
        st.session_state.dados["CARRETA_2"] = ""
        st.session_state.dados["CARRETA_3"] = ""

    if st.button("Avançar ➡️"):
        if all([
            st.session_state.dados['PLACA_CAMINHAO'],
            st.session_state.dados['KM_ATUAL'],
            st.session_state.dados['MOTORISTA']
        ]):
            st.session_state.dados["DATA"] = datetime.now().strftime("%d/%m/%Y")
            st.session_state.dados["HORA"] = datetime.now().strftime("%H:%M")
            st.session_state.etapa = 2
        else:
            st.warning("Preencha todos os campos obrigatórios.")

# -----------------
# ETAPA 2
# -----------------
elif st.session_state.etapa == 2:
    st.subheader("Inserção das Imagens.")
    st.image("Checklist.png", caption="Exemplo dos ângulos corretos para as fotos", use_container_width=True)   
    
    imagens = st.file_uploader("Envie ao menos 4 fotos", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
    if imagens and len(imagens) >= 4:
        st.session_state.imagens = imagens
        if st.button("Avançar ➡️"):
            st.session_state.etapa = 3
    else:
        st.warning("Envie no mínimo 4 imagens.")

# -----------------
# -----------------
# -----------------
# -----------------
# -----------------
elif st.session_state.etapa == 3:
    import time  # usado apenas para breve atraso visual antes do rerun

    st.subheader("Etapa 3: Checklist")
    checklist_itens = {
        "ARREFECIMENTO_OK": "Nível do líquido de arrefecimento",
        "OLEO_MOTOR_OK": "Nível de óleo de motor",
        "VAZAMENTO_OLEO_MOTOR": "Vazamento de óleo motor",
        "VAZAMENTO_AGUA_MOTOR": "Vazamento de água motor",
        "OLEO_CAMBIO_OK": "Vazamento de óleo câmbio",
        "OLEO_DIFERENCIAL_OK": "Vazamento de óleo diferencial",
        "OLEO_CUBOS_OK": "Vazamento de óleo cubos",
        "DIESEL_OK": "Vazamento de diesel",
        "GNV_OK": "Vazamento de GNV",
        "VAZAMENTO_AR_OK": "Vazamento de ar",
        "PNEUS_OK": "Pneus avariados",
        "FAIXAS_REFLETIVAS_OK": "Faixas refletivas",
        "FUNILARIA_OK": "Itens avariados para funilaria",
        "ILUMINACAO_OK": "Iluminação",
        "PARABRISA_OK": "Para-brisa",
        "FALHAS_PAINEL_OK": "Presença de falhas no painel",
        "TACOGRAFO_OK": "Funcionamento tacógrafo",
        "CÂMERA_PARABRISA": "Câmera do para-brisa",
        "CÂMERA_COLUNALD": "Câmera Coluna Lado Direito",
        "CÂMERA_COLUNALE": "Câmera Coluna Lado Esquerdo",
        "CÂMERA_DEFLETORLD": "Câmera Defletor Lado Direito",
        "CÂMERA_DEFLETORLE": "Câmera Defletor Lado Esquerdo",
        "PORTAL_OK": "Imagem Digital",
        "FUNCIONAMENTO_TK_OK": "Funcionamento TK"
    }

    for chave, descricao in checklist_itens.items():
        opcao = st.radio(
            descricao,
            options=["OK", "NÃO OK"],
            index=0,
            key=f"radio_{chave}",
            horizontal=True
        )
        st.session_state.dados[chave] = opcao
        if opcao == "NÃO OK":
            foto = st.file_uploader(f"Foto de {descricao}", type=['jpg', 'jpeg', 'png'], key=f"foto_{chave}")
            if foto:
                st.session_state.fotos_nao_ok[chave] = foto

    # Garantir inicialização de finalizando
    if "finalizando" not in st.session_state:
        st.session_state.finalizando = False

    # Botão para finalizar checklist
    if st.button("✅ Finalizar Checklist", disabled=st.session_state.finalizando):
        st.session_state.finalizando = True
        with st.spinner("Finalizando checklist..."):
            try:
                # Gerar Word
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

                # Gerar PDF
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

                # Enviar e-mail
                if enviar_email(buffer_word, buffer_pdf, st.session_state.fotos_nao_ok):
                    st.success("Checklist concluído com Sucesso! Reiniciando...")
                    time.sleep(1.5)  # pequena pausa para exibir mensagem
                    st.session_state.clear()
                    st.session_state.etapa = 1
                    st.rerun()
                else:
                    st.session_state.finalizando = False
                    st.error("O checklist foi gerado, mas o envio do e-mail falhou.")
                    st.stop()

            except Exception as e:
                st.session_state.finalizando = False
                st.error(f"Erro ao finalizar checklist: {e}")
