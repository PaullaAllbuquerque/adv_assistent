import os
import time
import json
import imaplib
import email
import logging
from email.header import decode_header
from pathlib import Path
from typing import List, Optional
from dotenv import load_dotenv
from google import genai
from google.genai.errors import APIError
import pandas as pd
from pydantic import BaseModel, Field
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google_drive import upload_planilha_drive


# Carrega o .env da raiz do projeto
caminho_raiz = Path(__file__).resolve().parent.parent / '.env'
load_dotenv(dotenv_path=caminho_raiz)

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# Diretórios do projeto
DIRETORIO_RAIZ = Path(__file__).resolve().parent.parent
DIRETORIO_DADOS = DIRETORIO_RAIZ / 'data'
DIRETORIO_DADOS.mkdir(exist_ok=True)
CAMINHO_PLANILHA = DIRETORIO_DADOS / "Acompanhamento_Demandas.xlsx"

DIRETORIO_LOGS = DIRETORIO_RAIZ / 'logs'
DIRETORIO_LOGS.mkdir(exist_ok=True)
CAMINHO_LOG = DIRETORIO_LOGS / 'sistema.log'

# Configuração do Sistema de Logs (Caixa-Preta / Auditoria)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(CAMINHO_LOG, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# Inicializa o cliente do Gemini
client = genai.Client(api_key=GEMINI_API_KEY)

# ==========================================
# MODELO PYDANTIC (Validação de Dados)
# ==========================================
class AudienciaVinicius(BaseModel):
    id: Optional[int] = Field(default=None, description="ID sequencial numérico")
    data_inicial: str = Field(alias="Data Inicial", description="Data da audiência no formato DD/MM/AAAA")
    hora_inicial: str = Field(alias="Hora Inicial", description="Horário da audiência HH:MM")
    compromisso: str = Field(alias="Compromisso", description="Tipo de audiência")
    local: str = Field(alias="Local", description="Localidade ou cidade")
    distribuicao_atual: str = Field(alias="Distribuição Atual", description="Vara do Trabalho")
    cliente: str = Field(alias="Cliente", description="Nome do cliente/empresa")
    processo_partes: str = Field(alias="Processo / Partes", description="Partes e número unificado do processo")
    numero_pasta: str = Field(alias="Número da Pasta", description="Número da pasta")
    encarregado: str = Field(alias="Encarregado", description="Sempre VINICIUS")
    materia: str = Field(alias="Matéria", description="Matéria do processo")
    anotacoes_links: str = Field(alias="Anotações / Links", description="Links de Zoom, IDs, senhas e despachos")

class ListaAudiencias(BaseModel):
    audiencias: List[AudienciaVinicius]


def extrair_corpo_email(msg):
    """Extrai o texto e/ou HTML do corpo do e-mail."""
    corpo = ""
    if msg.is_multipart():
        for parte in msg.walk():
            content_type = parte.get_content_type()
            content_disposition = str(parte.get("Content-Disposition", ""))
            if content_type in ["text/plain", "text/html"] and "attachment" not in content_disposition:
                payload = parte.get_payload(decode=True)
                if payload:
                    corpo += payload.decode(errors="ignore") + "\n"
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            corpo = payload.decode(errors="ignore")
    return corpo


def extrair_audiencias_vinicius_com_gemini(corpo_email):
    logging.info("Enviando o corpo do e-mail para o Gemini extrair as audiências do VINICIUS (com Pydantic)...")
    
    prompt = f"""
    Você é um Assistente Jurídico Inteligente especializado em pautas trabalhistas.
    Abaixo está o conteúdo de um e-mail contendo uma pauta semanal de audiências.

    SUA TAREFA PRINCIPAL:
    1. Identifique todas as linhas da tabela de audiências.
    2. Filtre e retorne APENAS as audiências em que o Encarregado é "VINICIUS" (ou "VINÍCIUS"). Ignore audiências de outros advogados.
    3. Retorne o resultado ESTRITAMENTE em formato JSON contendo uma chave "audiencias" que seja uma lista de objetos com os campos exatos solicitados.

    Campos obrigatórios por objeto:
    - "Data Inicial"
    - "Hora Inicial"
    - "Compromisso"
    - "Local"
    - "Distribuição Atual"
    - "Cliente"
    - "Processo / Partes"
    - "Número da Pasta"
    - "Encarregado" (deve ser "VINICIUS")
    - "Matéria"
    - "Anotações / Links"

    Conteúdo do E-mail:
    {corpo_email}
    """

    tentativas = 3
    for tentativa in range(tentativas):
        try:
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
                config={
                    "response_mime_type": "application/json"
                }
            )
            
            dados_brutos = json.loads(response.text)
            if isinstance(dados_brutos, list):
                dados_brutos = {"audiencias": dados_brutos}
                
            validado = ListaAudiencias(**dados_brutos)
            logging.info(f"Sucesso na extração e validação via Pydantic. Total de {len(validado.audiencias)} audiência(s) do Vinícius.")
            return [item.model_dump(by_alias=True) for item in validado.audiencias]

        except APIError as e:
            if e.code == 503 and tentativa < tentativas - 1:
                logging.warning(f"Servidor do Gemini instável (503). Tentando novamente em 5 segundos... (Tentativa {tentativa + 1}/{tentativas})")
                time.sleep(5)
            else:
                logging.error(f"Erro crítico na API do Gemini: {e}")
                raise e
        except Exception as e:
            logging.error(f"Erro na validação do Pydantic ou JSON: {e}")
            if tentativa == tentativas - 1:
                return []
            time.sleep(2)


def aplicar_layout_branding_vinicius(caminho_arquivo):
    """Aplica o layout executivo de forma segura na planilha Excel."""
    wb = openpyxl.load_workbook(caminho_arquivo)
    ws = wb.active

    # Usamos uma fonte em negrito limpa para o cabeçalho
    fonte_cabecalho = Font(name="Calibri", size=11, bold=True)
    
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    fonte_celulas = Font(name="Calibri", size=10)
    alinhamento_central = Alignment(horizontal="center", vertical="center")
    alinhamento_esquerda = Alignment(horizontal="left", vertical="center")

    # Formata o cabeçalho real da tabela na linha 1
    linha_cabecalho = 1
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=linha_cabecalho, column=col)
        cell.font = fonte_cabecalho
        cell.alignment = alinhamento_central

    # Formata as linhas de dados (a partir da linha 2)
    for row in range(linha_cabecalho + 1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = fonte_celulas
            cell.border = borda_fina
            
            if col in [1, 2, 3, 10]:
                cell.alignment = alinhamento_central
            else:
                cell.alignment = alinhamento_esquerda

    # Ajusta largura das colunas automaticamente
    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(caminho_arquivo)
    logging.info("Layout executivo aplicado com sucesso à planilha Excel.")
# Sincroniza automaticamente com o Google Drive do escritório

    logging.info("A iniciar o upload da planilha para o Google Drive...")
    upload_planilha_drive()

def atualizar_ou_criar_planilha(novas_audiencias):
    if not novas_audiencias:
        logging.info("Nenhuma audiência do Vinícius validada para adicionar à planilha.")
        return

    logging.info("Atualizando a planilha de acompanhamento personalizada do Vinícius...")
    df_novos = pd.DataFrame(novas_audiencias)
    
    colunas_ordenadas = [
        "ID", "Data Inicial", "Hora Inicial", "Compromisso", "Local", 
        "Distribuição Atual", "Cliente", "Processo / Partes", 
        "Número da Pasta", "Encarregado", "Matéria", "Anotações / Links"
    ]

    if CAMINHO_PLANILHA.exists():
        df_antigo = pd.read_excel(CAMINHO_PLANILHA)
        
        if "ID" not in df_antigo.columns:
            df_antigo["ID"] = range(1, len(df_antigo) + 1)

        colunas_chave = ["Data Inicial", "Hora Inicial", "Processo / Partes"]
        
        df_concat = pd.concat([df_antigo, df_novos], ignore_index=True)
        df_limpo = df_concat.drop_duplicates(subset=colunas_chave, keep='last').copy()

        df_limpo["ID"] = range(1, len(df_limpo) + 1)
        df_final = df_limpo
    else:
        df_novos["ID"] = range(1, len(df_novos) + 1)
        df_final = df_novos

    colunas_presentes = [c for c in colunas_ordenadas if c in df_final.columns]
    df_final = df_final[colunas_presentes]

    df_final.to_excel(CAMINHO_PLANILHA, index=False)
    aplicar_layout_branding_vinicius(CAMINHO_PLANILHA)
    
    logging.info(f"Planilha de Vinícius Bismarque formatada e salva com sucesso em: {CAMINHO_PLANILHA}")


def conectar_e_ler_emails():
    if not EMAIL_USER or not EMAIL_PASS:
        logging.critical("Erro crítico: Credenciais de e-mail não encontradas no .env!")
        return

    try:
        logging.info("Conectando ao servidor IMAP do Gmail...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        remetente_alvo = "paulamlleal@gmail.com"
        logging.info(f"Filtrando e-mails do remetente: {remetente_alvo}")
        
        status, messages = mail.search(None, f'(FROM "{remetente_alvo}")')
        if status != "OK" or not messages[0]:
            logging.warning("Nenhum e-mail encontrado para este remetente.")
            return

        email_ids = messages[0].split()
        ultimos_ids = email_ids[-1:] 
        
        logging.info("Processando o ÚLTIMO e-mail recebido com branding oficial...")
        
        for e_id in ultimos_ids:
            res, msg_data = mail.fetch(e_id, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    assunto, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(assunto, bytes):
                        assunto = assunto.decode(encoding if encoding else "utf-8", errors="ignore")
                    
                    remetente = msg["From"]
                    logging.info(f"E-mail Lido -> De: {remetente} | Assunto: {assunto}")
                    
                    corpo_email = extrair_corpo_email(msg)
                    
                    if corpo_email.strip():
                        audiencias_vinicius = extrair_audiencias_vinicius_com_gemini(corpo_email)
                        atualizar_ou_criar_planilha(audiencias_vinicius)
                    else:
                        logging.warning("Aviso: O corpo do e-mail veio vazio.")
                    
        mail.logout()
        logging.info("Conexão IMAP encerrada com sucesso.")

    except Exception as e:
        logging.error(f"Erro crítico ao processar o fluxo: {e}")

if __name__ == "__main__":
    conectar_e_ler_emails()