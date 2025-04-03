import os
import re
import pandas as pd
import requests
from openpyxl import load_workbook
from bs4 import BeautifulSoup

EXCEL_FILE = "acervo.xlsx"

def carregar_todo_excel():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    return pd.DataFrame(columns=[
        "Título", 
        "Autor", 
        "Editora", 
        "Ano", 
        "ISBN", 
        "Sigla", 
        "Número do Exemplar"
    ])

def salvar_em_excel(df):
    df.to_excel(EXCEL_FILE, index=False)

def obter_dados_livro(isbn):
    url = f"https://brasilapi.com.br/api/isbn/v1/{isbn}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            data = response.json()
            return {
                "Título": data.get("title", ""),
                "Autor": ", ".join(data.get("authors", [])),
                "Editora": data.get("publisher", ""),
                "Ano": data.get("year", ""),
                "ISBN": isbn
            }
    except Exception:
        pass
    return None

def sincronizar_dados(comando):
    padrao = re.compile(
        r"(?:Cadastrar|Registrar)?\s*(\d+)\s*(?:exemplar(?:es)?)?\s*(?:da\s+obra)?\s*(?:com)?\s*ISBN[:\s]*([\d\-xX]+)\s*(?:e\s*sigla\s*(?:da biblioteca)?[:\s]*([A-Z]{2,10}))",
        re.IGNORECASE
    )
    match = padrao.search(comando)

    if not match:
        return {
            "erro": "Comando inválido. Use: Cadastrar X exemplares da obra com ISBN: Y e sigla da biblioteca: Z"
        }

    quantidade, isbn, sigla = match.groups()
    quantidade = int(quantidade)
    isbn = isbn.replace("-", "").strip().upper()

    dados_livro = obter_dados_livro(isbn)
    if not dados_livro:
        return {
            "erro": "Erro ao processar o comando: ISBN não encontrado na BrasilAPI."
        }

    acervo = carregar_todo_excel()
    ultimo_numero = acervo[acervo["ISBN"] == isbn]["Número do Exemplar"].max()
    if pd.isna(ultimo_numero):
        ultimo_numero = 0

    novos_dados = []
    for i in range(1, quantidade + 1):
        novo = dados_livro.copy()
        novo["Sigla"] = sigla.upper()
        novo["Número do Exemplar"] = int(ultimo_numero + i)
        novos_dados.append(novo)

    acervo = pd.concat([acervo, pd.DataFrame(novos_dados)], ignore_index=True)
    salvar_em_excel(acervo)

    return {
        "mensagem": f"{quantidade} exemplar(es) do livro '{dados_livro['Título']}' foram cadastrados com sucesso para a biblioteca '{sigla.upper()}'.",
        "dados": novos_dados
    }

# ✅ Compatível com cadastro.py
def sincronizar_acervo(comando):
    return sincronizar_dados(comando)
