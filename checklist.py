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
import zipfile
import time

# Carregar variáveis de ambiente
load_dotenv()

# Configuração da página
st.set_page_config(page_title="Checklist de Caminhão", layout="centered")
st.title("📝 CheckList Manutenção")

# Estados
if "etapa" not in st.session_state:
    st.session_state.etapa = 1
if "dados" not in st.session_state:
    st.session_state.dados = {}
if "imagens" not in st.session_state:
    st.session_state.imagens = []
if "fotos_nao_ok" not in st.session_state:
    st.session_state.fotos_nao_ok = {}

# -------------------
# RESPONSÁVEIS POR ITEM (grupos)
# -------------------
RESPONSAVEIS = {
    (
        "ruberval.silva@transmaroni.com.br",
        "alex.franca@transmaroni.com.br",
        "jose.oliveira@transmaroni.com.br",
        "sarah.ferreira@transmaroni.com.br",
        "enielle.argolo@transmaroni.com.br",
        "michele.silva@transmaroni.com.br",
        "sabrina.silva@transmaroni.com.br"
    ): [
        "VAZAMENTO_OLEO_MOTOR", "VAZAMENTO_AGUA_MOTOR", "OLEO_MOTOR_OK", "ARREFECIMENTO_OK",
        "OLEO_CAMBIO_OK", "OLEO_DIFERENCIAL_OK", "DIESEL_OK", "GNV_OK", "OLEO_CUBOS_OK",
        "VAZAMENTO_AR_OK", "PNEUS_OK", "PARABRISA_OK", "ILUMINACAO_OK", "FAIXAS_REFLETIVAS_OK",
        "FALHAS_PAINEL_OK"
    ],
    ("lucas.alves@transmaroni.com.br",): [
        "FUNCIONAMENTO_TK_OK"
    ],
    ("sandra.silva@transmaroni.com.br", "amanda.soares@transmaroni.com.br", "lucas.alves@transmaroni.com.br"): [
        "TACOGRAFO_OK"
    ],
    ("wesley.assumpcao@transmaroni.com.br", "bruna.silva@transmaroni.com.br", "alex.franca@transmaroni.com.br"): [
        "FUNILARIA_OK"
    ],
    # Grupo de câmeras e imagem digital
    ("mirella.trindade@transmaroni.com.br",): [
        "CÂMERA_COLUNALD", "CÂMERA_COLUNALE", "CÂMERA_DEFLETORLD", "CÂMERA_DEFLETORLE",
        "CÂMERA_PARABRISA",
        # Itens IMAGEM DIGITAL solicitados (devem existir no DOCX como {{...}})
        "CÂMERACOLUNA_LD", "CÂMERACOLUNA_LE", "CÂMERADEFLETOR_LD", "CÂMERADEFLETOR_LE"
    ],
}

# -------------------
# FUNÇÕES AUXILIARES
# -------------------
def gerar_zip_imagens(imagens):
    """Cria um ZIP com as imagens da etapa 2"""
    buffer_zip = BytesIO()
    with zipfile.ZipFile(buffer_zip, "w") as zf:
        for idx, img in enumerate(imagens, start=1):
            zf.writestr(f"foto_{idx}.jpg", img.getvalue())
    buffer_zip.seek(0)
    return buffer_zip

def enviar_emails_personalizados(itens_nao_ok, fotos_nao_ok, checklist_itens, buffer_word, buffer_zip):
    """Envia os e-mails para os responsáveis de cada item com Word, ZIP e fotos dos itens"""
    hora_atual = datetime.now().hour
    saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"

    for destinatarios, itens_responsaveis in RESPONSAVEIS.items():
        itens_do_grupo = [i for i in itens_nao_ok if i in itens_responsaveis]
        if not itens_do_grupo:
            continue

        msg = EmailMessage()
        msg["Subject"] = f" CHECKLIST DE MANUTENÇÃO - {st.session_state.dados.get('PLACA_CAMINHAO','')}"
        msg["From"] = os.getenv("EMAIL_USER")
        msg["To"] = ", ".join(destinatarios)

        itens_texto = "\n".join([f"- {checklist_itens[i]}" for i in itens_do_grupo])
        msg.set_content(
            f"{saudacao},\n\n"
            f"Motorista: {st.session_state.dados.get('MOTORISTA','')}\n"
            f"Vistoriador: {st.session_state.dados.get('VISTORIADOR','')}\n"
            f"Data: {st.session_state.dados.get('DATA','')} {st.session_state.dados.get('HORA','')}\n\n"
            f"O veículo {st.session_state.dados.get('PLACA_CAMINHAO','')} foi verificado em seu CHECKLIST.\n"
            f"Os seguintes itens foram vistoriados e precisam ser encaminhados para manutenção:\n\n"
            f"{itens_texto}\n\n"
            "Atenciosamente,\nSistema de Checklist"
        )

        # Anexar Ficha Técnica (Word)
        msg.add_attachment(
            buffer_word.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="Ficha_Tecnica.docx"
        )

        # Anexar ZIP das fotos da etapa 2
        msg.add_attachment(
            buffer_zip.getvalue(),
            maintype="application",
            subtype="zip",
            filename="Fotos_Checklist.zip"
        )

        # Anexar fotos dos itens NÃO OK (somente os do grupo)
        for item in itens_do_grupo:
            if item in fotos_nao_ok:
                arquivos = fotos_nao_ok[item]
                if not isinstance(arquivos, list):
                    arquivos = [arquivos]
                for idx, foto in enumerate(arquivos, start=1):
                    msg.add_attachment(
                        foto.getvalue(),
                        maintype="image",
                        subtype="jpeg",
                        filename=f"{item}_{idx}.jpg"
                    )

        try:
            with smtplib.SMTP(os.getenv("EMAIL_HOST"), int(os.getenv("EMAIL_PORT"))) as smtp:
                smtp.starttls()
                smtp.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS"))
                smtp.send_message(msg)
        except Exception as e:
            st.error(f"Erro ao enviar e-mail para {destinatarios}: {e}")

# === Integração com SharePoint ===
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

def gerar_payload_sharepoint(dados_checklist):
    """Monta o payload a ser enviado para a lista do SharePoint"""
    km_str = str(dados_checklist.get("KM_ATUAL", "0")).replace(".", "").replace(",", "")
    try:
        km_int = int(km_str)
    except ValueError:
        km_int = 0
    return {
        "Title": dados_checklist.get("PLACA_CAMINHAO", ""),
        "field_0": datetime.now().isoformat(),
        "field_2": dados_checklist.get("PLACA_CAMINHAO", ""),
        "field_3": dados_checklist.get("PLACA_CARRETA2", ""),
        "field_4": dados_checklist.get("MOTORISTA", ""),
        "field_5": dados_checklist.get("VISTORIADOR", ""),
        "field_6": km_int,
        "field_7": dados_checklist.get("TIPO_VEICULO", ""),
        "field_8": dados_checklist.get("CARRETA_2", ""),
        "field_9": dados_checklist.get("OBSERVACOES", "")
    }

def enviar_para_sharepoint(payload):
    """Envia os dados do checklist para a lista do SharePoint"""
    site_url = "https://grupotransmaroni.sharepoint.com/sites/MCDMANUTENOCORPORATIVADIGITAL"
    username = "lucas.alves@transmaroni.com.br"
    password = "Maroni@23"
    list_name = "SBD Checklist Manutenção Oficial"

    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
        target_list = ctx.web.lists.get_by_title(list_name)
        item = target_list.add_item(payload)
        ctx.execute_query()
    else:
        raise Exception("Falha na autenticação com o SharePoint")

# -------------------
# ETAPA 1
# -------------------
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

# -------------------
# ETAPA 2
# -------------------
elif st.session_state.etapa == 2:
    st.subheader("Inserção das Imagens")
    st.image("Checklist.png", caption="Exemplo dos ângulos corretos", use_container_width=True)

    imagens = st.file_uploader(
        "Envie ao menos 4 fotos",
        type=['jpg', 'jpeg', 'png'],
        accept_multiple_files=True
    )
    if imagens and len(imagens) >= 4:
        st.session_state.imagens = imagens

    col1, col2 = st.columns(2)
    if col1.button("⬅️ Voltar"):
        st.session_state.etapa = 1
        st.rerun()
    if col2.button("Avançar ➡️"):
        if st.session_state.imagens and len(st.session_state.imagens) >= 4:
            st.session_state.etapa = 3
        else:
            st.warning("Envie no mínimo 4 imagens.")
# -------------------
# ETAPA 3
# -------------------
elif st.session_state.etapa == 3:
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
        "FUNCIONAMENTO_TK_OK": "Funcionamento TK",

        # ===== Itens IMAGEM DIGITAL solicitados (chaves EXATAS do DOCX) =====
        "CÂMERACOLUNA_LD": "Imagem digital câmera coluna LD",
        "CÂMERACOLUNA_LE": "Imagem digital câmera coluna LE",
        "CÂMERADEFLETOR_LD": "Imagem digital câmera defletor LD",
        "CÂMERADEFLETOR_LE": "Imagem digital câmera defletor LE",
    }

    for chave, descricao in checklist_itens.items():
        opcao = st.radio(
            descricao,
            ["OK", "NÃO OK"],
            index=0,
            key=f"radio_{chave}",
            horizontal=True
        )
        st.session_state.dados[chave] = opcao
        if opcao == "NÃO OK":
            fotos = st.file_uploader(
                f"Fotos de {descricao}",
                type=['jpg', 'jpeg', 'png'],
                key=f"foto_{chave}",
                accept_multiple_files=True
            )
            if fotos:
                st.session_state.fotos_nao_ok[chave] = fotos

    # Campo OBSERVAÇÕES solicitado
    st.session_state.dados["OBSERVACOES"] = st.text_area("Observações", placeholder="Digite informações adicionais, se necessário.")

    if "finalizando" not in st.session_state:
        st.session_state.finalizando = False

    col1, col2 = st.columns(2)
    if col1.button("⬅️ Voltar"):
        st.session_state.etapa = 2
        st.rerun()

    if col2.button("✅ Finalizar Checklist", disabled=st.session_state.finalizando):
        st.session_state.finalizando = True
        with st.spinner("Finalizando checklist..."):
            try:
                # ===== Gera Word a partir do template =====
                # Certifique-se que o arquivo do template corresponde ao seu DOCX:
                # Ex.: "Ficha Técnica.docx" e contém os placeholders {{CHAVE}}
                doc = Document("Ficha Técnica.docx")
                for p in doc.paragraphs:
                    for k, v in st.session_state.dados.items():
                        token = f"{{{{{k}}}}}"
                        if token in p.text:
                            p.text = p.text.replace(token, str(v))
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                for k, v in st.session_state.dados.items():
                                    token = f"{{{{{k}}}}}"
                                    if token in p.text:
                                        p.text = p.text.replace(token, str(v))

                buffer_word = BytesIO()
                doc.save(buffer_word)
                buffer_word.seek(0)

                # (Mantido) Gera PDF simples com os dados - opcional
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

                # ZIP das fotos etapa 2
                buffer_zip = gerar_zip_imagens(st.session_state.imagens)

                # Itens NÃO OK (para e-mails)
                itens_nao_ok = [k for k, v in st.session_state.dados.items() if v == "NÃO OK"]

                # Envia e-mails com Word, ZIP e fotos dos itens
                enviar_emails_personalizados(
                    itens_nao_ok,
                    st.session_state.fotos_nao_ok,
                    checklist_itens,
                    buffer_word,
                    buffer_zip
                )

                # ===== Envia para o SharePoint =====
                try:
                    payload = gerar_payload_sharepoint(st.session_state.dados)
                    enviar_para_sharepoint(payload)
                    st.success("Checklist concluído, e-mails enviados e dados enviados ao SharePoint! Reiniciando...")
                except Exception as e:
                    st.warning(f"Checklist concluído e e-mails enviados, mas NÃO foi salvo no SharePoint: {e}")

                time.sleep(2)
                st.session_state.clear()
                st.session_state.etapa = 1
                st.rerun()

            except Exception as e:
                st.session_state.finalizando = False
                st.error(f"Erro ao finalizar checklist: {e}")
