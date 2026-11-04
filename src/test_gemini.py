import os
from dotenv import load_dotenv
from google import genai

# Carrega as variáveis do arquivo .env (onde está sua chave)
load_dotenv()

# Inicializa o cliente do Gemini
client = genai.Client()

try:
    # Faz um teste simples de conexão enviando uma mensagem para a IA
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents="Responda apenas com a frase: 'Conexão com o Gemini estabelecida com sucesso!'"
    )
    print("\n[SUCESSO]:", response.text)
except Exception as e:
    print(f"\n[ERRO]: Ocorreu um erro ao conectar com o Gemini: {e}")