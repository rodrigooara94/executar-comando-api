import re
import requests
import pandas as pd
import os

ARQUIVO_EXCEL = 'acervo.xlsx'

def carregar_todo_excel():
    if os.path.exists(ARQUIVO_EXCEL):
        return pd.read_excel(ARQUIVO_EXCEL)
    else:
        return pd.DataFrame(columns=['Título', 'Autor', 'ISBN', 'Editora', 'Ano', 'Exemplar', 'Sigla'])

def salvar_em_excel(df):
    df.to_excel(ARQUIVO_EXCEL, index=False)

def sincronizar_acervo():
    df = carregar_todo_excel()
    df = df.drop_duplicates()
    salvar_em_excel(df)
    return f"{len(df)} registros sincronizados com sucesso!"

def buscar_dados_por_isbn(isbn):
    url = f"https://brasilapi.com.br/api/isbn/v1/{isbn}"
    response = requests.get(url)

    if response.status_code == 200:
        return response.json()
    else:
        return None

def sincronizar_dados(comando):
    try:
        # Regex flexível para extrair os dados (quantidade, ISBN e sigla)
        match = re.search(
            r'(\d+)\s+exemplares?.*?ISBN[:\s]*([\d\-xX]+).*?sigla[:\s]*([\w-]+)',
            comando,
            re.IGNORECASE
        )

        if not match:
            return {"erro": "Comando inválido. Use algo como: Cadastrar X exemplares da obra com ISBN: Y e sigla: Z"}

        quantidade = int(match.group(1))
        isbn = match.group(2).replace('-', '')
        sigla = match.group(3).upper()

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
                'Sigla': sigla
            }
            df = pd.concat([df, pd.DataFrame([novo_exemplar])], ignore_index=True)

        salvar_em_excel(df)

        return {
            "mensagem": f"{quantidade} exemplares da obra '{dados_livro.get('title')}' com ISBN {isbn} e sigla {sigla} foram sincronizados com sucesso!"
        }

    except Exception as e:
        return {"erro": f"Erro ao processar o comando: {str(e)}"}
