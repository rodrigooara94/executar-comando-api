import pandas as pd
import os
import re
import requests
from openpyxl import load_workbook

CAMINHO_ARQUIVO_EXCEL = 'acervo.xlsx'
NOME_PLANILHA = 'Acervo'

def carregar_todo_excel():
    if not os.path.exists(CAMINHO_ARQUIVO_EXCEL):
        colunas = ['Título', 'Autor', 'ISBN', 'Editora', 'Ano', 'Exemplar', 'Sigla']
        df = pd.DataFrame(columns=colunas)
        salvar_em_excel(df)
    else:
        df = pd.read_excel(CAMINHO_ARQUIVO_EXCEL, sheet_name=NOME_PLANILHA)
    return df

def salvar_em_excel(df):
    with pd.ExcelWriter(CAMINHO_ARQUIVO_EXCEL, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=NOME_PLANILHA, index=False)

def buscar_dados_por_isbn(isbn):
    try:
        url = f"https://brasilapi.com.br/api/isbn/v1/{isbn}"
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.json()
        else:
            return None
    except Exception:
        return None

def sincronizar_dados(comando):
    try:
        # Regex mais permissiva com suporte a "sigla da biblioteca"
        regex = re.compile(
            r'(?P<quantidade>\d+).*?isbn[:\s-]*?(?P<isbn>[\d\-xX]+).*?(sigla( da biblioteca)?[:\s-]*)?(?P<sigla>[\w-]+)',
            re.IGNORECASE
        )
        match = regex.search(comando)

        if not match:
            return {
                "erro": (
                    "Comando inválido. Exemplo de uso flexível:\n"
                    "'Cadastrar 3 exemplares da obra com ISBN 9781234567890 e sigla da biblioteca XYZ'"
                )
            }

        quantidade = int(match.group('quantidade'))
        isbn = match.group('isbn').replace('-', '').strip()
        sigla = match.group('sigla').strip().upper()

        dados_livro = buscar_dados_por_isbn(isbn)
        if not dados_livro:
            return {"erro": "Erro ao processar o comando: ISBN não encontrado na BrasilAPI."}

        df = carregar_todo_excel()

        for i in range(quantidade):
            novo_exemplar = {
                'Título': dados_livro.get('title', ''),
                'Autor': ', '.join(dados_livro.get('authors', [])),
                'ISBN': isbn,
                'Editora': dados_livro.get('publisher', ''),
                'Ano': dados_livro.get('year', ''),
                'Exemplar': i + 1,
                'Sigla': sigla  # Internamente continua como 'Sigla'
            }
            df = pd.concat([df, pd.DataFrame([novo_exemplar])], ignore_index=True)

        salvar_em_excel(df)

        return {
            "mensagem": (
                f"{quantidade} exemplares da obra '{dados_livro.get('title')}' com ISBN {isbn} "
                f"foram sincronizados com sucesso para a biblioteca com sigla {sigla}!"
            )
        }

    except Exception as e:
        return {"erro": f"Erro ao processar o comando: {str(e)}"}
