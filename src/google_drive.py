import os
import logging
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Escopos necessários para criar e gerir ficheiros no Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def autenticar_google_drive():
    """Autentica o aplicativo utilizando o credentials.json e gera o token.json."""
    creds = None
    
    # O token.json armazena os tokens de acesso e atualização do utilizador
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    # Se não houver credenciais válidas, permite que o utilizador faça login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                raise FileNotFoundError("O ficheiro 'credentials.json' não foi encontrado na raiz do projeto!")
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
            
        # Salva as credenciais para as próximas execuções
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return creds

def upload_planilha_drive(caminho_arquivo_local="data/Acompanhamento_Demandas.xlsx", nome_no_drive="Acompanhamento_Demandas_Vinicius.xlsx"):
    """Faz o upload ou atualização da planilha de audiências diretamente para o Google Drive."""
    try:
        creds = autenticar_google_drive()
        service = build('drive', 'v3', credentials=creds)
        
        # Procura se o ficheiro já existe no Drive para atualizar em vez de duplicar
        query = f"name='{nome_no_drive}' and trashed=false"
        results = service.files().list(q=query, spaces='drive', fields="files(id, name)").execute()
        files = results.get('files', [])
        
        media = MediaFileUpload(caminho_arquivo_local, resumable=True)
        
        if files:
            # Atualiza o ficheiro existente
            file_id = files[0]['id']
            updated_file = service.files().update(
                fileId=file_id,
                media_body=media
            ).execute()
            logging.info(f"Planilha atualizada com sucesso no Google Drive! ID: {updated_file.get('id')}")
        else:
            # Cria um novo ficheiro no Drive
            file_metadata = {'name': nome_no_drive}
            uploaded_file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            logging.info(f"Nova planilha enviada com sucesso para o Google Drive! ID: {uploaded_file.get('id')}")
            
    except Exception as e:
        logging.error(f"Erro ao sincronizar com o Google Drive: {e}")
        raise e