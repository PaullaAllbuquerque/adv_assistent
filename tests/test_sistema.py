import pytest
import pandas as pd
from pydantic import ValidationError
import sys
from pathlib import Path

# Adiciona a pasta src ao path para conseguir importar o script principal se necessário
sys.path.append(str(Path(__file__).resolve().parent.parent / 'src'))

from ler_email import AudienciaVinicius

def test_validacao_pydantic_audiencia_valida():
    """Testa se o Pydantic aceita uma audiência com os campos corretos do Vinícius."""
    dados_validos = {
        "Data Inicial": "03/11/2026",
        "Hora Inicial": "08:05",
        "Compromisso": "Audiência inicial",
        "Local": "Betim",
        "Distribuição Atual": "1ª Vara do Trabalho de Betim",
        "Cliente": "CIMCOP S/A",
        "Processo / Partes": "GERALDO X CIMCOP",
        "Número da Pasta": "002.751",
        "Encarregado": "VINICIUS",
        "Matéria": "RESPONSABILIDADE SUBSIDIÁRIA",
        "Anotações / Links": "Link do Zoom: https://zoom.us/j/123"
    }
    
    audiencia = AudienciaVinicius(**dados_validos)
    assert audiencia.encarregado == "VINICIUS"
    assert audiencia.data_inicial == "03/11/2026"


def test_validacao_pydantic_campos_faltando():
    """Testa se o Pydantic rejeita dados incompletos."""
    dados_incompletos = {
        "Data Inicial": "03/11/2026"
        # Faltam os outros campos obrigatórios
    }
    
    with pytest.raises(ValidationError):
        AudienciaVinicius(**dados_incompletos)


def test_logica_deduplicacao_pandas():
    """Testa se o Pandas remove linhas duplicadas baseando-se na chave única."""
    # Simula dados antigos na planilha
    df_antigo = pd.DataFrame([
        {
            "ID": 1,
            "Data Inicial": "03/11/2026",
            "Hora Inicial": "08:05",
            "Processo / Partes": "GERALDO X CIMCOP",
            "Encarregado": "VINICIUS",
            "Anotações / Links": "Link antigo do Zoom"
        }
    ])
    
    # Simula dados novos vindos do e-mail (mesma data, hora e processo, mas com link do Zoom levemente alterado)
    df_novo = pd.DataFrame([
        {
            "ID": None,
            "Data Inicial": "03/11/2026",
            "Hora Inicial": "08:05",
            "Processo / Partes": "GERALDO X CIMCOP",
            "Encarregado": "VINICIUS",
            "Anotações / Links": "Link novo do Zoom atualizado"
        }
    ])
    
    colunas_chave = ["Data Inicial", "Hora Inicial", "Processo / Partes"]
    df_concat = pd.concat([df_antigo, df_novo], ignore_index=True)
    df_limpo = df_concat.drop_duplicates(subset=colunas_chave, keep='last').copy()
    
    # A regra de ouro: deve continuar tendo apenas 1 linha (evitou a duplicata e manteve a última)
    assert len(df_limpo) == 1
    assert df_limpo.iloc[0]["Anotações / Links"] == "Link novo do Zoom atualizado"