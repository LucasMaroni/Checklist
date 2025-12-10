import streamlit as st
import os
from datetime import datetime
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
import smtplib
from email.message import EmailMessage
from email.header import Header
from email.utils import formataddr
from dotenv import load_dotenv
import zipfile
import time
import pandas as pd
import re

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

# Mapeamento de e-mails das operações
EMAILS_OPERACOES = {
    "MERCADO - LIVRE": ["meli.operacional@transmaroni.com.br", "programacaoecommerce@transmaroni.com.br", "lucas.alves@transmaroni.com.br"],
    "BITREM": ["bitremgrupo@transmaroni.com.br"],
    "FRIGO": ["frigogrupo@transmaroni.com.br"],
    "BIMBO": ["adm.bimbo@transmaroni.com.br"],
    "BAÚ": ["baugrupo@transmaroni.com.br"]
}

def formatar_remetente():
    """Formata o remetente de forma profissional"""
    return formataddr((
        str(Header('Sistema de Checklist - TransMaroni', 'utf-8')),
        os.getenv("EMAIL_USER")
    ))

def criar_corpo_email_html(dados, checklist_itens, itens_nao_ok):
    """Cria corpo do e-mail em HTML para melhor aceitação"""
    hora_atual = datetime.now().hour
    saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
    
    # Criar listas de itens OK e NÃO OK
    itens_ok_lista = []
    itens_nao_ok_lista = []
    
    for chave, descricao in checklist_itens.items():
        status = dados.get(chave, "N/A")
        if status == "OK":
            itens_ok_lista.append(f"<li style='color: green;'>✓ {descricao}</li>")
        else:
            itens_nao_ok_lista.append(f"<li style='color: red;'>✗ {descricao}</li>")
    
    itens_ok_html = "".join(itens_ok_lista) if itens_ok_lista else "<li>Nenhum item com problema</li>"
    itens_nao_ok_html = "".join(itens_nao_ok_lista) if itens_nao_ok_lista else "<li>Todos os itens estão OK</li>"
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            body {{
                font-family: 'Arial', sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                background-color: #f9f9f9;
            }}
            .header {{
                background-color: #0047AB;
                color: white;
                padding: 20px;
                border-radius: 10px 10px 0 0;
                text-align: center;
            }}
            .content {{
                background-color: white;
                padding: 30px;
                border-radius: 0 0 10px 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .section {{
                margin-bottom: 25px;
                padding-bottom: 15px;
                border-bottom: 1px solid #eee;
            }}
            .section-title {{
                color: #0047AB;
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 10px;
                padding-bottom: 5px;
                border-bottom: 2px solid #0047AB;
            }}
            .info-row {{
                display: flex;
                margin-bottom: 8px;
            }}
            .info-label {{
                font-weight: bold;
                width: 200px;
                color: #555;
            }}
            .info-value {{
                flex: 1;
            }}
            .status-ok {{
                color: #28a745;
                font-weight: bold;
            }}
            .status-nok {{
                color: #dc3545;
                font-weight: bold;
            }}
            .footer {{
                margin-top: 30px;
                padding-top: 20px;
                border-top: 1px solid #ddd;
                font-size: 12px;
                color: #666;
                text-align: center;
            }}
            .disclaimer {{
                background-color: #f8f9fa;
                padding: 15px;
                border-left: 4px solid #0047AB;
                margin: 20px 0;
                font-size: 13px;
                color: #6c757d;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1 style="margin: 0; font-size: 24px;">🚚 CHECKLIST DE MANUTENÇÃO</h1>
            <p style="margin: 5px 0 0 0; opacity: 0.9;">Sistema Oficial - TransMaroni Logística</p>
        </div>
        
        <div class="content">
            <div class="disclaimer">
                <strong>⚠️ E-MAIL OFICIAL DO SISTEMA:</strong> Esta é uma mensagem automática do sistema oficial de checklist da TransMaroni. 
                Você está recebendo este e-mail porque é responsável por itens de manutenção verificados.
            </div>
            
            <div class="section">
                <div class="section-title">📋 DADOS DA VISTORIA</div>
                <div class="info-row">
                    <div class="info-label">Motorista:</div>
                    <div class="info-value">{dados.get('MOTORISTA', 'N/A')}</div>
                </div>
                <div class="info-row">
                    <div class="info-label">Vistoriador:</div>
                    <div class="info-value">{dados.get('VISTORIADOR', 'N/A')}</div>
                </div>
                <div class="info-row">
                    <div class="info-label">Data/Hora:</div>
                    <div class="info-value">{dados.get('DATA', 'N/A')} {dados.get('HORA', 'N/A')}</div>
                </div>
            </div>
            
            <div class="section">
                <div class="section-title">🚛 DADOS DO VEÍCULO</div>
                <div class="info-row">
                    <div class="info-label">Placa Caminhão:</div>
                    <div class="info-value"><strong>{dados.get('PLACA_CAMINHAO', 'N/A')}</strong></div>
                </div>
                <div class="info-row">
                    <div class="info-label">KM Atual:</div>
                    <div class="info-value">{dados.get('KM_ATUAL', 'N/A')}</div>
                </div>
                <div class="info-row">
                    <div class="info-label">Operação:</div>
                    <div class="info-value">{dados.get('OPERACAO', 'N/A')}</div>
                </div>
                <div class="info-row">
                    <div class="info-label">Tipo de Veículo:</div>
                    <div class="info-value">{dados.get('TIPO_VEICULO', 'N/A')}</div>
                </div>
            </div>
            
            <div class="section">
                <div class="section-title">✅ RESUMO DA VISTORIA</div>
                <div class="info-row">
                    <div class="info-label">Total Itens OK:</div>
                    <div class="info-value status-ok">{len(itens_ok_lista)} itens</div>
                </div>
                <div class="info-row">
                    <div class="info-label">Total Itens NÃO OK:</div>
                    <div class="info-value status-nok">{len(itens_nao_ok_lista)} itens</div>
                </div>
            </div>
            
            <div class="section">
                <div class="section-title">🔧 ITENS COM PROBLEMAS</div>
                <ul>
                    {itens_nao_ok_html}
                </ul>
            </div>
            
            <div class="section">
                <div class="section-title">📝 OBSERVAÇÕES</div>
                <p>{dados.get('OBSERVACOES', 'Nenhuma observação registrada.')}</p>
            </div>
            
            <div class="footer">
                <p><strong>Sistema de Checklist Automático</strong><br>
                TransMaroni Logística<br>
                Este e-mail foi gerado automaticamente. Por favor, não responda.<br>
                Para suporte técnico: manutencao.frota@transmaroni.com.br</p>
                <p style="font-size: 10px; color: #999;">
                    ID: {datetime.now().strftime('%Y%m%d%H%M%S')} | 
                    Tempo execução: {calcular_tempo_execucao()}
                </p>
            </div>
        </div>
    </body>
    </html>
    """
    return html

def enviar_email_seguro(smtp_host, smtp_port, username, password, msg):
    """Envia e-mail com fallback SSL/TLS para melhor compatibilidade"""
    try:
        # Primeira tentativa: SSL (porta 465)
        with smtplib.SMTP_SSL(smtp_host, 465, timeout=30) as smtp:
            smtp.login(username, password)
            smtp.send_message(msg)
        return True, "SSL"
    except Exception as e:
        st.warning(f"SSL falhou ({e}), tentando STARTTLS...")
        try:
            # Segunda tentativa: STARTTLS (porta 587)
            with smtplib.SMTP(smtp_host, smtp_port, timeout=30) as smtp:
                smtp.starttls()
                smtp.login(username, password)
                smtp.send_message(msg)
            return True, "STARTTLS"
        except Exception as e2:
            return False, f"STARTTLS também falhou: {e2}"

def enviar_emails_personalizados(itens_nao_ok, fotos_nao_ok, checklist_itens, buffer_word, buffer_zip):
    """Envia os e-mails para os responsáveis de cada item com Word, ZIP e fotos dos itens"""
    
    # Configurações SMTP do .env
    smtp_host = os.getenv("EMAIL_HOST")
    smtp_port = int(os.getenv("EMAIL_PORT", 587))
    username = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    
    if not all([smtp_host, username, password]):
        st.error("Configurações de e-mail não encontradas no .env")
        return

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
        msg["Subject"] = f"[OFICIAL] CHECKLIST DE MANUTENÇÃO - {st.session_state.dados.get('PLACA_CAMINHAO','')}"
        msg["From"] = formatar_remetente()
        msg["To"] = ", ".join(todos_destinatarios)
        
        # Headers importantes para evitar spam
        msg["X-Mailer"] = "TransMaroni Checklist System v2.0"
        msg["X-Auto-Response-Suppress"] = "OOF, AutoReply"
        msg["X-Priority"] = "3"
        msg["X-MSMail-Priority"] = "Normal"
        msg["Importance"] = "Normal"
        msg["List-Unsubscribe"] = f"<mailto:{username}?subject=Unsubscribe>"

        # Corpo do e-mail em HTML e texto plano
        texto_plano = f"""
        CHECKLIST DE MANUTENÇÃO - TransMaroni
        ======================================
        
        Motorista: {st.session_state.dados.get('MOTORISTA','')}
        Vistoriador: {st.session_state.dados.get('VISTORIADOR','')}
        Data: {st.session_state.dados.get('DATA','')} {st.session_state.dados.get('HORA','')}
        Placa: {st.session_state.dados.get('PLACA_CAMINHAO','')}
        
        O veículo {st.session_state.dados.get('PLACA_CAMINHAO','')} foi verificado em seu CHECKLIST.
        Itens que precisam de manutenção:
        
        {chr(10).join([f'- {checklist_itens[i]}' for i in itens_do_grupo])}
        
        ---
        Este é um e-mail automático do sistema oficial de checklist da TransMaroni.
        Para suporte: manutencao.frota@transmaroni.com.br
        """
        
        # Corpo HTML
        html_content = criar_corpo_email_html(
            st.session_state.dados, 
            checklist_itens, 
            itens_do_grupo
        )
        
        msg.set_content(texto_plano)
        msg.add_alternative(html_content, subtype='html')

        # Anexar Ficha Técnica (Word)
        msg.add_attachment(
            buffer_word.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"Ficha_Tecnica_{st.session_state.dados.get('PLACA_CAMINHAO','')}.docx"
        )

        # Anexar ZIP das fotos da etapa 2
        msg.add_attachment(
            buffer_zip.getvalue(),
            maintype="application",
            subtype="zip",
            filename=f"Fotos_Checklist_{st.session_state.dados.get('PLACA_CAMINHAO','')}.zip"
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
            sucesso, metodo = enviar_email_seguro(smtp_host, smtp_port, username, password, msg)
            if sucesso:
                st.success(f"E-mail enviado para {len(todos_destinatarios)} destinatários ({metodo})")
            else:
                st.error(f"Falha ao enviar e-mail para {todos_destinatarios}: {metodo}")
        except Exception as e:
            st.error(f"Erro ao enviar e-mail para {todos_destinatarios}: {e}")

def enviar_email_lucas(checklist_itens, buffer_word, buffer_zip, fotos_nao_ok):
    """Envia e-mail completo para Lucas Alves com todos os dados do checklist em formato estruturado"""
    
    smtp_host = os.getenv("EMAIL_HOST")
    smtp_port = int(os.getenv("EMAIL_PORT", 587))
    username = os.getenv("EMAIL_USER")
    password = os.getenv("EMAIL_PASS")
    
    if not all([smtp_host, username, password]):
        st.error("Configurações de e-mail não encontradas para envio ao Lucas")
        return False
    
    hora_atual = datetime.now().hour
    saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
    
    msg = EmailMessage()
    msg["Subject"] = f"[RELATÓRIO COMPLETO] CHECKLIST {st.session_state.dados.get('PLACA_CAMINHAO','')}"
    msg["From"] = formatar_remetente()
    msg["To"] = "lucas.alves@transmaroni.com.br"
    
    # Headers para melhor aceitação
    msg["X-Report-Type"] = "Checklist Completo"
    msg["X-Vehicle-Plate"] = st.session_state.dados.get('PLACA_CAMINHAO', '')
    
    # Criar resumo dos itens
    itens_ok = []
    itens_nao_ok = []
    
    for chave, descricao in checklist_itens.items():
        status = st.session_state.dados.get(chave, "N/A")
        if status == "OK":
            itens_ok.append(f"✓ {descricao}")
        else:
            itens_nao_ok.append(f"✗ {descricao}")
    
    # Corpo do e-mail formatado para facilitar extração com Power Automate
    corpo_email = f"""
    {saudacao}, Lucas
    
    ========== RELATÓRIO COMPLETO - SISTEMA DE CHECKLIST ==========
    
    [DADOS DO VEÍCULO]
    PLACA: {st.session_state.dados.get('PLACA_CAMINHAO', 'N/A')}
    KM_ATUAL: {st.session_state.dados.get('KM_ATUAL', 'N/A')}
    MOTORISTA: {st.session_state.dados.get('MOTORISTA', 'N/A')}
    OPERACAO: {st.session_state.dados.get('OPERACAO', 'N/A')}
    VISTORIADOR: {st.session_state.dados.get('VISTORIADOR', 'N/A')}
    TIPO_VEICULO: {st.session_state.dados.get('TIPO_VEICULO', 'N/A')}
    DATA: {st.session_state.dados.get('DATA', 'N/A')}
    HORA: {st.session_state.dados.get('HORA', 'N/A')}
    PLACA_CARRETA1: {st.session_state.dados.get('PLACA_CARRETA1', 'N/A')}
    PLACA_CARRETA2: {st.session_state.dados.get('PLACA_CARRETA2', 'N/A')}
    BITREM: {st.session_state.dados.get('BITREM', 'N/A')}
    
    [ESTATÍSTICAS]
    ITENS_OK: {len(itens_ok)}
    ITENS_NOK: {len(itens_nao_ok)}
    TOTAL_ITENS: {len(checklist_itens)}
    PERCENTUAL_OK: {round(len(itens_ok)/len(checklist_itens)*100, 1)}%
    
    [ITENS EM CONFORMIDADE]
    {chr(10).join(itens_ok) if itens_ok else "Nenhum item com problema"}
    
    [ITENS COM PROBLEMAS]
    {chr(10).join(itens_nao_ok) if itens_nao_ok else "Todos os itens estão OK"}
    
    [OBSERVAÇÕES]
    {st.session_state.dados.get('OBSERVACOES', 'Nenhuma observação registrada.')}
    
    [METADADOS]
    TIMESTAMP: {datetime.now().isoformat()}
    TEMPO_EXECUCAO: {calcular_tempo_execucao()}
    SISTEMA_VERSION: 2.0
    EMAIL_ID: {datetime.now().strftime('%Y%m%d%H%M%S')}
    
    ---
    Este é um relatório automático do Sistema de Checklist da TransMaroni.
    Dúvidas: manutencao.frota@transmaroni.com.br
    """
    
    msg.set_content(corpo_email)
    
    # Anexar Ficha Técnica (Word)
    msg.add_attachment(
        buffer_word.getvalue(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=f"Ficha_Tecnica_{st.session_state.dados.get('PLACA_CAMINHAO','')}.docx"
    )
    
    # Anexar ZIP das fotos da etapa 2
    msg.add_attachment(
        buffer_zip.getvalue(),
        maintype="application",
        subtype="zip",
        filename=f"Fotos_Checklist_{st.session_state.dados.get('PLACA_CAMINHAO','')}.zip"
    )
    
    # Anexar fotos dos itens NÃO OK
    for item, fotos in fotos_nao_ok.items():
        if fotos:
            arquivos = fotos if isinstance(fotos, list) else [fotos]
            for idx, foto in enumerate(arquivos, start=1):
                msg.add_attachment(
                    foto.getvalue(),
                    maintype="image",
                    subtype="jpeg",
                    filename=f"{item}_{idx}.jpg"
                )
    
    try:
        sucesso, metodo = enviar_email_seguro(smtp_host, smtp_port, username, password, msg)
        if sucesso:
            st.success(f"E-mail para Lucas Alves enviado com sucesso ({metodo})")
            return True
        else:
            st.error(f"Falha ao enviar e-mail para Lucas Alves: {metodo}")
            return False
    except Exception as e:
        st.error(f"Erro ao enviar e-mail para Lucas Alves: {e}")
        return False

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
        
        # Verificar configurações de e-mail
        if not all([os.getenv("EMAIL_HOST"), os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASS")]):
            st.error("⚠️ Configurações de e-mail não encontradas no .env!")
            st.info("""
            Configure as seguintes variáveis no arquivo `.env`:
            ```
            EMAIL_HOST=smtp.office365.com
            EMAIL_PORT=587
            EMAIL_USER=seu-email@transmaroni.com.br
            EMAIL_PASS=sua-senha
            ```
            """)
            st.session_state.finalizando = False
            st.stop()
        
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

                # Exibir resumo
                st.info(f"""
                **Resumo do Checklist:**
                - Itens OK: {len(checklist_itens) - len(itens_nao_ok)}
                - Itens NÃO OK: {len(itens_nao_ok)}
                - Total de fotos: {len(st.session_state.imagens)}
                - Destinatários: Vários grupos de e-mail
                """)

                # Envia e-mails para os responsáveis pelos itens
                with st.spinner("Enviando e-mails para os responsáveis..."):
                    enviar_emails_personalizados(
                        itens_nao_ok,
                        st.session_state.fotos_nao_ok,
                        checklist_itens,
                        buffer_word,
                        buffer_zip
                    )
                
                # ENVIO DO E-MAIL PARA O LUCAS
                with st.spinner("Enviando relatório completo para Lucas Alves..."):
                    enviou_lucas = enviar_email_lucas(checklist_itens, buffer_word, buffer_zip, st.session_state.fotos_nao_ok)
                
                if enviou_lucas:
                    st.success("✅ Checklist concluído e e-mail enviado para Lucas Alves!")
                    
                    # Mostrar dicas para evitar problemas de spam
                    with st.expander("💡 Dicas para garantir recebimento dos e-mails"):
                        st.markdown("""
                        1. **Peça ao Lucas para:**
                           - Adicionar `manutencao.frota@transmaroni.com.br` aos contatos
                           - Criar regra no Outlook para mover seus e-mails para a Caixa de Entrada
                        
                        2. **Configurações do domínio (TI da TransMaroni):**
                           - Configurar SPF, DKIM e DMARC no domínio `transmaroni.com.br`
                           - Usar subdomínio específico: `sistema@checklist.transmaroni.com.br`
                        
                        3. **Verificação:**
                           - Testar em https://mxtoolbox.com/blacklists.aspx
                           - Usar https://www.mail-tester.com/ para pontuação de spam
                        """)
                else:
                    st.warning("⚠️ Checklist concluído, mas houve problema com o e-mail para Lucas Alves.")

                time.sleep(3)
                st.session_state.clear()
                st.session_state.etapa = 1
                st.rerun()

            except Exception as e:
                st.session_state.finalizando = False
                st.error(f"❌ Erro ao finalizar checklist: {e}")
                st.exception(e)
