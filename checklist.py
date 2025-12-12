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
import zipfile
import time
import pandas as pd
import re
from msal import ConfidentialClientApplication
import requests
import json

@st.cache_resource
def carregar_placas_validas():
    try:
        df_placas = pd.read_excel("placas.xlsx")
        return set(df_placas['PLACA'].str.upper().str.strip())
    except Exception as e:
        st.warning(f"Não foi possível carregar a lista de placas: {e}")
        return set()

# Carregar lista de placas válidas ANTES de qualquer uso
PLACAS_VALIDAS = set()
try:
    df_placas = pd.read_excel("placas.xlsx")
    PLACAS_VALIDAS = set(df_placas['PLACA'].str.upper().str.strip())
except Exception as e:
    st.warning(f"Não foi possível carregar a lista de placas: {e}")

# Carregar variáveis de ambiente
load_dotenv()

# Configurações do SharePoint (adicione essas variáveis no seu .env)
SHAREPOINT_SITE_URL = "https://grupotransmaroni.sharepoint.com/sites/ProTech"
SHAREPOINT_LIST_NAME = "SGBD - Checklists de Manutenção 📝"
TENANT_ID = "8babaf2b-248f-4651-98d6-1e5312fbe00a"
CLIENT_ID = "82db6f97-74e8-458a-9273-c8aa3e1c7e7a"
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "vlM8Q~k~IvWd1rYMePChe4T6S0gaInWEakWSsazM")

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
        "manutencao.frota@transmaroni.com.br",
        "sabrina.silva@transmaroni.com.br"
    ): [
        "VAZAMENTO_OLEO_MOTOR", "VAZAMENTO_AGUA_MOTOR", "OLEO_MOTOR_OK", "ARREFECIMENTO_OK",
        "OLEO_CAMBIO_OK", "OLEO_DIFERENCIAL_OK", "DIESEL_OK", "GNV_OK", "OLEO_CUBOS_OK",
        "VAZAMENTO_AR_OK", "PNEUS_OK", "PARABRISA_OK", "ILUMINACAO_OK", "FAIXAS_REFLETIVAS_OK",
        "FALHAS_PAINEL_OK"
    ],
    ("henrique.araujo@transmaroni.com.br", "amanda.soares@transmaroni.com.br", "manutencao.frota@transmaroni.com.br",): [
        "FUNCIONAMENTO_TK_OK"
    ],
    ("sandra.silva@transmaroni.com.br", "amanda.soares@transmaroni.com.br", "manutencao.frota@transmaroni.com.br", ): [
        "TACOGRAFO_OK"
    ],
    ("wesley.assumpcao@transmaroni.com.br", "manutencao.frota@transmaroni.com.br", "bruna.silva@transmaroni.com.br", "alex.franca@transmaroni.com.br", ): [
        "FUNILARIA_OK"
    ],
    # Grupo de câmeras e imagem digital
    ("mirella.trindade@transmaroni.com.br", "manutencao.frota@transmaroni.com.br", ): [
        "CÂMERA_COLUNALD", "CÂMERA_COLUNALE", "CÂMERA_DEFLETORLD", "CÂMERA_DEFLETORLE",
        "CÂMERA_PARABRISA",
        # Itens IMAGEM DIGITAL solicitados (devem existir no DOCX como {{...}})
        "CÂMERACOLUNA_LD", "CÂMERACOLUNA_LE", "CÂMERADEFLETOR_LD", "CÂMERADEFLETOR_LE"
    ],
    (
        "manutencao.frota@transmaroni.com.br",
        "ruberval.silva@transmaroni.com.br",
        "michele.silva@transmaroni.com.br",
        "enielle.argolo@transmaroni.com.br",
        "jose.oliveira@transmaroni.com.br",
        "eric.souza@transmaroni.com.br",
    ): [
        "PARAFUSO_SUSPENSAO_VANDERLEIA_FACCHINI"
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

def obter_token_sharepoint():
    """Obtém token de acesso para SharePoint usando Client Secret"""
    try:
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        app = ConfidentialClientApplication(
            client_id=CLIENT_ID,
            client_credential=CLIENT_SECRET,
            authority=authority
        )
        
        # Escopo para SharePoint Online
        scopes = ["https://graph.microsoft.com/.default"]
        
        result = app.acquire_token_for_client(scopes=scopes)
        
        if "access_token" in result:
            return result["access_token"]
        else:
            st.error(f"Erro ao obter token: {result.get('error_description')}")
            return None
    except Exception as e:
        st.error(f"Erro na autenticação: {e}")
        return None

def enviar_para_sharepoint(dados, fotos_nao_ok):
    """Envia dados do checklist para lista do SharePoint"""
    try:
        # Obter token de acesso
        access_token = obter_token_sharepoint()
        if not access_token:
            return False
        
        # URL da API Graph para a lista
        list_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_URL}:/lists/{SHAREPOINT_LIST_NAME}/items"
        
        # Mapeamento dos campos do checklist para os field_N do SharePoint
        item_data = {
            "fields": {
                "Title": dados.get("PLACA_CAMINHAO", ""),  # field_0/Title
                "field_1": dados.get("OPERACAO", ""),      # OPERAÇÃO
                "field_2": f"{dados.get('DATA', '')} {dados.get('HORA', '')}",  # DATA E HORA
                "field_3": dados.get("PLACA_CAMINHAO", ""),  # PLACA CAMINHÃO 1
                "field_4": dados.get("PLACA_CARRETA2", ""),  # PLACA CAMINHÃO 2 (BITREM?)
                "field_5": calcular_tempo_execucao(),       # TIME EXECUÇÃO
                "field_6": dados.get("MOTORISTA", ""),      # MOTORISTA
                "field_7": dados.get("VISTORIADOR", ""),    # VISTORIADOR
                "field_8": dados.get("KM_ATUAL", ""),       # KM ATUAL
                "field_9": dados.get("TIPO_VEICULO", ""),   # TIPO DE VEÍCULO
                "field_10": "2 EIXOS" if dados.get("CARRETA_2") == "X" else "3 EIXOS" if dados.get("CARRETA_3") == "X" else "",  # TIPO DE CARRETA
                "field_11": dados.get("OBSERVACOES", ""),   # OBSERVAÇÃO
                
                # Checklist - mapeamento dos itens técnicos
                "field_12": dados.get("ARREFECIMENTO_OK", "N/A"),  # NÍVEL DO LÍQUIDO DE ARREFECIMENTO
                "field_13": dados.get("OLEO_MOTOR_OK", "N/A"),     # NÍVEL DE ÓLEO DE MOTOR
                "field_14": dados.get("VAZAMENTO_OLEO_MOTOR", "N/A"),  # VAZAMENTO DE ÓLEO MOTOR
                "field_15": dados.get("VAZAMENTO_AGUA_MOTOR", "N/A"),  # VAZAMENTO DE ÁGUA MOTOR
                "field_16": dados.get("OLEO_CAMBIO_OK", "N/A"),    # VAZAMENTO DE ÓLEO CÂMBIO
                "field_17": dados.get("OLEO_DIFERENCIAL_OK", "N/A"),  # VAZAMENTO DE ÓLEO DIFERENCIAL
                "field_18": dados.get("OLEO_CUBOS_OK", "N/A"),     # VAZAMENTO DE ÓLEO CUBOS
                "field_19": dados.get("DIESEL_OK", "N/A"),         # VAZAMENTO DE DIESEL
                "field_20": dados.get("GNV_OK", "N/A"),            # VAZAMENTO DE GNV
                "field_21": dados.get("VAZAMENTO_AR_OK", "N/A"),   # VAZAMENTO DE AR
                "field_22": dados.get("PNEUS_OK", "N/A"),          # PNEUS AVARIADOS
                "field_23": dados.get("FAIXAS_REFLETIVAS_OK", "N/A"),  # FAIXAS REFLETIVAS
                "field_24": dados.get("FUNILARIA_OK", "N/A"),      # ITENS AVARIADOS PARA FUNILÁRIA
                "field_25": dados.get("ILUMINACAO_OK", "N/A"),     # ILUMINAÇÃO
                "field_26": dados.get("PARABRISA_OK", "N/A"),      # PARA-BRISA
                "field_27": dados.get("FALHAS_PAINEL_OK", "N/A"),  # PRESENÇA DE FALHAS NO PAINEL
                "field_28": dados.get("TACOGRAFO_OK", "N/A"),      # FUNCIONAMENTO TACÓGRAFO
                "field_29": dados.get("CÂMERA_PARABRISA", "N/A"),  # CÂMERA DO PARABRISA
                "field_30": dados.get("CÂMERA_COLUNALD", "N/A"),   # CÂMERA COLUNA LADO DIREITO
                "field_31": dados.get("CÂMERA_COLUNALE", "N/A"),   # CÂMERA COLUNA LADO ESQUERDO
                "field_32": dados.get("CÂMERA_DEFLETORLD", "N/A"), # CÂMERA DEFLETOR LADO DIREITO
                "field_33": dados.get("CÂMERA_DEFLETORLE", "N/A"), # CÂMERA DEFLETOR LADO ESQUERDO
                "field_34": dados.get("FUNCIONAMENTO_TK_OK", "N/A"),  # FUNCIONAMENTO TK
                "field_35": dados.get("CÂMERACOLUNA_LD", "N/A"),   # IMAGEM DIGITAL CÂMERA COLUNA LD
                "field_36": dados.get("CÂMERACOLUNA_LE", "N/A"),   # IMAGEM DIGITAL CÂMERA COLUNA LE
                "field_37": dados.get("CÂMERADEFLETOR_LD", "N/A"), # IMAGEM DIGITAL CÂMERA DEFLETOR LD
                "field_38": dados.get("CÂMERADEFLETOR_LE", "N/A"), # IMAGEM DIGITAL CÂMERA DEFLETOR LE
                "field_39": dados.get("PARAFUSO_SUSPENSAO_VANDERLEIA_FACCHINI", "N/A"),  # PARAFUSO SUSPENSÃO
            }
        }
        
        # Headers para a requisição
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        # Criar item na lista
        response = requests.post(list_url, headers=headers, json=item_data)
        
        if response.status_code == 201:
            item_id = response.json().get("id")
            st.success(f"Dados enviados para SharePoint! ID do item: {item_id}")
            
            # Enviar fotos como anexos
            if fotos_nao_ok:
                enviar_fotos_como_anexos(item_id, fotos_nao_ok, access_token)
            
            return True
        else:
            st.error(f"Erro ao enviar para SharePoint: {response.status_code} - {response.text}")
            return False
            
    except Exception as e:
        st.error(f"Erro ao enviar para SharePoint: {e}")
        return False

def enviar_fotos_como_anexos(item_id, fotos_nao_ok, access_token):
    """Envia fotos dos itens NÃO OK como anexos do item do SharePoint"""
    try:
        for item_nome, fotos in fotos_nao_ok.items():
            if fotos:
                arquivos = fotos if isinstance(fotos, list) else [fotos]
                for idx, foto in enumerate(arquivos, start=1):
                    # URL para upload do anexo
                    upload_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_URL}:/lists/{SHAREPOINT_LIST_NAME}/items/{item_id}/attachments"
                    
                    # Criar o conteúdo do arquivo
                    file_content = foto.getvalue()
                    file_name = f"{item_nome}_{idx}.jpg"
                    
                    # Headers para upload
                    headers = {
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/octet-stream"
                    }
                    
                    # Fazer upload do arquivo
                    response = requests.post(
                        upload_url,
                        headers=headers,
                        data=file_content,
                        params={"fileName": file_name}
                    )
                    
                    if response.status_code != 201:
                        st.warning(f"Erro ao enviar anexo {file_name}: {response.status_code}")
                        
    except Exception as e:
        st.warning(f"Erro ao enviar fotos como anexos: {e}")

# Mapeamento de e-mails das operações
EMAILS_OPERACOES = {
    "MERCADO - LIVRE": ["meli.operacional@transmaroni.com.br", "programacaoecommerce@transmaroni.com.br", "lucas.alves@transmaroni.com.br"],
    "BITREM": ["bitremgrupo@transmaroni.com.br"],
    "FRIGO": ["frigogrupo@transmaroni.com.br"],
    "BIMBO": ["adm.bimbo@transmaroni.com.br"],
    "BAÚ": ["baugrupo@transmaroni.com.br"]
}

def enviar_emails_personalizados(itens_nao_ok, fotos_nao_ok, checklist_itens, buffer_word, buffer_zip):
    """Envia os e-mails para os responsáveis de cada item com Word, ZIP e fotos dos itens"""
    hora_atual = datetime.now().hour
    saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"

    # Adiciona e-mails das operações conforme selecionado
    operacao = st.session_state.dados.get("OPERACAO", "")
    emails_operacao = EMAILS_OPERACOES.get(operacao, [])

    for destinatarios, itens_responsaveis in RESPONSAVEIS.items():
        itens_do_grupo = [i for i in itens_nao_ok if i in itens_responsaveis]
        if not itens_do_grupo:
            continue

        # Junta os e-mails do grupo com os da operação (sem duplicar)
        todos_destinatarios = list(set(destinatarios + tuple(emails_operacao)))

        msg = EmailMessage()
        msg["Subject"] = f" CHECKLIST DE MANUTENÇÃO - {st.session_state.dados.get('PLACA_CAMINHAO','')}"
        msg["From"] = os.getenv("EMAIL_USER")
        msg["To"] = ", ".join(todos_destinatarios)

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
            st.error(f"Erro ao enviar e-mail para {todos_destinatarios}: {e}")

def calcular_tempo_execucao():
    """Calcula o tempo de execução do checklist"""
    if "start_time" in st.session_state:
        segundos = int(time.time() - st.session_state.start_time)
        minutos = segundos // 60
        segundos_restantes = segundos % 60
        return f"{minutos:02d}:{segundos_restantes:02d}"
    return "N/A"

# -------------------
# ETAPA 1
# -------------------
if st.session_state.etapa == 1:
    st.subheader("Dados do Veículo e Condutor")
    st.session_state.dados['PLACA_CAMINHAO'] = st.text_input("Placa do Caminhão", max_chars=7)

    # Inicia o temporizador quando o usuário começa a digitar
    if (
        "start_time" not in st.session_state
        and st.session_state.dados['PLACA_CAMINHAO']
    ):
        st.session_state.start_time = time.time()

    placa_digitada = st.session_state.dados['PLACA_CAMINHAO'].upper().strip()

    # Regex padrão Mercosul: LLLNLNN
    padrao_mercosul = r"^[A-Z]{3}[0-9][A-Z][0-9]{2}$"
    placa_padrao = bool(re.match(padrao_mercosul, placa_digitada))
    placa_valida = placa_digitada in PLACAS_VALIDAS if placa_digitada else False

    if placa_digitada and not placa_padrao:
        st.warning("Placa fora do padrão Mercosul! Use o formato LLLNLNN (ex: ABC1D23).")
    elif placa_digitada and not placa_valida:
        st.warning("PLACA INVÁLIDA! Verifique com sua equipe de manutenção.")

    st.session_state.dados['KM_ATUAL'] = st.text_input("KM Atual")
    st.session_state.dados['MOTORISTA'] = st.text_input("Motorista")

    # Campo OPERAÇÃO
    operacoes = [
        "MERCADO - LIVRE",
        "BITREM",
        "BIG",
        "CARREFOUR",
        "SOTREC",
        "FRIGO",
        "BIMBO",
        "UNILEVER",
        "BAÚ",
        "OUTROS",
        "PÁTIO"
    ]
    st.session_state.dados['OPERACAO'] = st.selectbox("Operação", operacoes)

    st.session_state.dados['VISTORIADOR'] = "ANTÔNIO RINALDO RAMOS"

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

    if st.button("Avançar ➡️", disabled=not (placa_padrao and placa_valida)):
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
        "CÂMERACOLUNA_LD": "Imagem digital câmera coluna LD",
        "CÂMERACOLUNA_LE": "Imagem digital câmera coluna LE",
        "CÂMERADEFLETOR_LD": "Imagem digital câmera defletor LD",
        "CÂMERADEFLETOR_LE": "Imagem digital câmera defletor LE",
        "PARAFUSO_SUSPENSAO_VANDERLEIA_FACCHINI": "Parafuso suspensão Vanderleia Facchini",
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

                # ZIP das fotos etapa 2
                buffer_zip = gerar_zip_imagens(st.session_state.imagens)

                # Itens NÃO OK (para e-mails)
                itens_nao_ok = [k for k, v in st.session_state.dados.items() if v == "NÃO OK"]

                # 1. Envia dados para SharePoint
                sucesso_sharepoint = enviar_para_sharepoint(
                    st.session_state.dados, 
                    st.session_state.fotos_nao_ok
                )
                
                # 2. Envia e-mails para os responsáveis pelos itens (mantém funcionalidade existente)
                enviar_emails_personalizados(
                    itens_nao_ok,
                    st.session_state.fotos_nao_ok,
                    checklist_itens,
                    buffer_word,
                    buffer_zip
                )
                
                if sucesso_sharepoint:
                    st.success("Checklist concluído! Dados enviados para SharePoint e e-mails enviados para os responsáveis.")
                else:
                    st.warning("Checklist concluído e e-mails enviados, mas houve problema com SharePoint.")

                time.sleep(2)
                st.session_state.clear()
                st.session_state.etapa = 1
                st.rerun()

            except Exception as e:
                st.session_state.finalizando = False
                st.error(f"Erro ao finalizar checklist: {e}")
