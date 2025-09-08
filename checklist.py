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
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
    df_placas = pd.read_excel("placas.xlsx")  # ajuste o nome do arquivo conforme necessário
    PLACAS_VALIDAS = set(df_placas['PLACA'].str.upper().str.strip())
except Exception as e:
    st.warning(f"Não foi possível carregar a lista de placas: {e}")

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
        "manutencao.frota@transmaroni.com.br",
        "sabrina.silva@transmaroni.com.br"
    ): [
        "VAZAMENTO_OLEO_MOTOR", "VAZAMENTO_AGUA_MOTOR", "OLEO_MOTOR_OK", "ARREFECIMENTO_OK",
        "OLEO_CAMBIO_OK", "OLEO_DIFERENCIAL_OK", "DIESEL_OK", "GNV_OK", "OLEO_CUBOS_OK",
        "VAZAMENTO_AR_OK", "PNEUS_OK", "PARABRISA_OK", "ILUMINACAO_OK", "FAIXAS_REFLETIVAS_OK",
        "FALHAS_PAINEL_OK"
    ],
    ("lucas.alves@transmaroni.com.br", "henrique.araujo@transmaroni.com.br", "amanda.soares@transmaroni.com.br", "manutencao.frota@transmaroni.com.br",): [
        "FUNCIONAMENTO_TK_OK"
    ],
    ("sandra.silva@transmaroni.com.br", "amanda.soares@transmaroni.com.br", "manutencao.frota@transmaroni.com.br", "lucas.alves@transmaroni.com.br", ): [
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

# Mapeamento dos itens para os nomes das colunas na planilha Excel
CHECKLIST_TO_EXCEL = {
    "VAZAMENTO_OLEO_MOTOR": "VAZAMENTO DE ÓLEO MOTOR",
    "VAZAMENTO_AGUA_MOTOR": "VAZAMENTO DE ÁGUA MOTOR",
    "OLEO_MOTOR_OK": "NÍVEL DE ÓLEO DE MOTOR",
    "ARREFECIMENTO_OK": "NÍVEL DO LÍQUIDO DE ARREFECIMENTO",
    "OLEO_CAMBIO_OK": "VAZAMENTO DE ÓLEO CÃMBIO",
    "OLEO_DIFERENCIAL_OK": "VAZAMENTO DE ÓLEO DIFERENCIAL",
    "DIESEL_OK": "VAZAMENTO DE DIESEL",
    "GNV_OK": "VAZAMENTO DE GNV",
    "OLEO_CUBOS_OK": "VAZAMENTO DE ÓLEO CUBOS",
    "VAZAMENTO_AR_OK": "VAZAMENTO DE AR",
    "PNEUS_OK": "PNEUS AVARIADOS",
    "PARABRISA_OK": "PARA-BRISA",
    "ILUMINACAO_OK": "ILUMINAÇÃO",
    "FAIXAS_REFLETIVAS_OK": "FAIXAS REFLETIVAS",
    "FALHAS_PAINEL_OK": "PRESENÇA DE FALHAS NO PAINEL",
    "FUNCIONAMENTO_TK_OK": "FUNCIONAMENTO TK",
    "TACOGRAFO_OK": "FUNCIONAMENTO TACÓGRAFO",
    "FUNILARIA_OK": "ITENS AVARIADOS PARA FUNILÁRIA",
    "CÂMERA_COLUNALD": "CÂMERA COLUNA LADO DIREITO",
    "CÂMERA_COLUNALE": "CÂMERA COLUNA LADO ESQUERDO",
    "CÂMERA_DEFLETORLD": "CÂMERA DEFLETOR LADO DIREITO",
    "CÂMERA_DEFLETORLE": "CÂMERA DEFLETOR LADO ESQUERDO",
    "CÂMERA_PARABRISA": "CÂMERA DO PARABRISA",
    "CÂMERACOLUNA_LD": "IMAGEM DIGITAL CÂMERA COLUNA LD",
    "CÂMERACOLUNA_LE": "IMAGEM DIGITAL CÂMERA COLUNA LE",
    "CÂMERADEFLETOR_LD": "IMAGEM DIGITAL CÂMERA DEFLETOR LD",
    "CÂMERADEFLETOR_LE": "IMAGEM DIGITAL CÂMERA DEFLETOR LE",
    "PARAFUSO_SUSPENSAO_VANDERLEIA_FACCHINI": "PARAFUSO SUSPENSÃO VANDERLEIA FACCHINI",
}

# Mapeamento das colunas da planilha Excel
EXCEL_COLUMNS = {
    "ID": "A",
    "Title": "B",
    "OPERAÇÃO": "C",
    "DATA E HORA": "D",
    "PLACA CAMINHÃO 1": "E",
    "PLACA CAMINHÃO 2": "F",
    "TIME EXECUÇÃO": "G",
    "MOTORISTA": "H",
    "VISTORIADOR": "I",
    "KM ATUAL": "J",
    "TIPO DE VEÍCULO": "K",
    "TIPO DE CARRETA": "L",
    "VAZAMENTO DE ÓLEO MOTOR": "M",
    "VAZAMENTO DE ÁGUA MOTOR": "N",
    "ITENS AVARIADOS PARA FUNILÁRIA": "O",
    "CÂMERA COLUNA LADO ESQUERDO": "P",
    "CÂMERA DEFLETOR LADO DIREITO": "Q",
    "CÂMERA COLUNA LADO DIREITO": "R",
    "NÍVEL DE ÓLEO DE MOTOR": "S",
    "FUNCIONAMENTO TACÓGRAFO": "T",
    "FUNCIONAMENTO TK": "U",
    "PRESENÇA DE FALHAS NO PAINEL": "V",
    "FAIXAS REFLETIVAS": "W",
    "ILUMINAÇÃO": "X",
    "PNEUS AVARIADOS": "Y",
    "PARA-BRISA": "Z",
    "VAZAMENTO DE AR": "AA",
    "VAZAMENTO DE ÓLEO DIFERENCIAL": "AB",
    "VAZAMENTO DE ÓLEO CUBOS": "AC",
    "VAZAMENTO DE GNV": "AD",
    "VAZAMENTO DE DIESEL": "AE",
    "CÂMERA DEFLETOR LADO ESQUERDO": "AF",
    "IMAGEM DIGITAL CÂMERA DEFLETOR LE": "AG",
    "IMAGEM DIGITAL CÂMERA DEFLETOR LD": "AH",
    "IMAGEM DIGITAL CÂMERA COLUNA LE": "AI",
    "IMAGEM DIGITAL CÂMERA COLUNA LD": "AJ",
    "CÂMERA DO PARABRISA": "AK",
    "PARAFUSO SUSPENSÃO VANDERLEIA FACCHINI": "AL",
    "NÍVEL DO LÍQUIDO DE ARREFECIMENTO": "AM",
    "VAZAMENTO DE ÓLEO CÃMBIO": "AN"
}

def salvar_na_planilha_excel(dados_checklist, tempo_execucao):
    """Salva os dados do checklist na planilha Excel"""
    try:
        # Carregar a planilha existente
        workbook = load_workbook("sbd-checklists.xlsx")
        sheet = workbook.active
        
        # Encontrar a próxima linha vazia
        next_row = sheet.max_row + 1
        
        # Gerar ID único (próximo número sequencial)
        sheet[f"A{next_row}"] = next_row - 1  # ID
        
        # Preencher os dados básicos
        sheet[f"B{next_row}"] = dados_checklist.get("PLACA_CAMINHAO", "").upper()  # Title
        sheet[f"C{next_row}"] = dados_checklist.get("OPERACAO", "")  # OPERAÇÃO
        sheet[f"D{next_row}"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # DATA E HORA
        sheet[f"E{next_row}"] = dados_checklist.get("PLACA_CAMINHAO", "").upper()  # PLACA CAMINHÃO 1
        sheet[f"F{next_row}"] = dados_checklist.get("PLACA_CARRETA2", "").upper()  # PLACA CAMINHÃO 2
        sheet[f"G{next_row}"] = tempo_execucao  # TIME EXECUÇÃO
        sheet[f"H{next_row}"] = dados_checklist.get("MOTORISTA", "").upper()  # MOTORISTA
        sheet[f"I{next_row}"] = dados_checklist.get("VISTORIADOR", "").upper()  # VISTORIADOR
        
        # KM ATUAL (converter para número)
        km_str = str(dados_checklist.get("KM_ATUAL", "0")).replace(".", "").replace(",", "")
        try:
            sheet[f"J{next_row}"] = int(km_str)
        except ValueError:
            sheet[f"J{next_row}"] = 0
        
        sheet[f"K{next_row}"] = dados_checklist.get("TIPO_VEICULO", "").upper()  # TIPO DE VEÍCULO
        
        # TIPO DE CARRETA
        if dados_checklist.get("CARRETA_2") == "X":
            sheet[f"M{next_row}"] = "2 EIXOS"
        elif dados_checklist.get("CARRETA_3") == "X":
            sheet[f"M{next_row}"] = "3 EIXOS"
        
        # Preencher os itens do checklist
        for chave, coluna_excel in CHECKLIST_TO_EXCEL.items():
            valor = dados_checklist.get(chave, "")
            if coluna_excel in EXCEL_COLUMNS:
                coluna = EXCEL_COLUMNS[coluna_excel]
                sheet[f"{coluna}{next_row}"] = valor.upper() if valor else ""
        
        # Salvar a planilha
        workbook.save("sbd-checklists.xlsx")
        return True
        
    except Exception as e:
        st.error(f"Erro ao salvar na planilha Excel: {e}")
        return False

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

    st.session_state.dados['VISTORIADOR'] = "CLEBER QUELSON BEZERRA DE MENEZES"

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

        # ===== Itens IMAGEM DIGITAL solicitados (chaves EXATAS do DOCX) =====
        "CÂMERACOLUNA_LD": "Imagem digital câmera coluna LD",
        "CÂMERACOLUNA_LE": "Imagem digital câmera coluna LE",
        "CÂMERADEFLETOR_LD": "Imagem digital câmera defletor LD",
        "CÂMERADEFLETOR_LE": "Imagem digital câmera defletor LE",
        "PARAFUSO_SUSPENSAO_VANDERLEIA_FACCHINI": "Parafuso suspensão Vanderleia Facchini",  # <-- Adicionado
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

                # Calcula o tempo de execução do checklist
                tempo_execucao = ""
                if "start_time" in st.session_state:
                    segundos = int(time.time() - st.session_state.start_time)
                    minutos = segundos // 60
                    segundos_restantes = segundos % 60
                    tempo_execucao = f"{minutos:02d}:{segundos_restantes:02d}"

                # ===== Salva na planilha Excel =====
                try:
                    sucesso = salvar_na_planilha_excel(st.session_state.dados, tempo_execucao)
                    if sucesso:
                        st.success("Checklist concluído, e-mails enviados e dados salvos na planilha Excel! Reiniciando...")
                    else:
                        st.warning("Checklist concluído e e-mails enviados, mas NÃO foi salvo na planilha Excel.")
                except Exception as e:
                    st.warning(f"Checklist concluído e e-mails enviados, mas NÃO foi salvo na planilha Excel: {e}")

                time.sleep(2)
                st.session_state.clear()
                st.session_state.etapa = 1
                st.rerun()

            except Exception as e:
                st.session_state.finalizando = False
                st.error(f"Erro ao finalizar checklist: {e}")
