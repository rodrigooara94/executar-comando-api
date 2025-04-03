import pandas as pd
import re
import unicodedata
import json
import requests
import shutil
import os
import requests

from gerar_docx import gerar_documento
from sincronizar import sincronizar_dados

def carregar_dicionario_cutter():
    caminho = "cutter_dict.json"
    if not os.path.exists(caminho):
        return {}
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_dicionario_cutter(dicionario):
    with open("cutter_dict.json", "w", encoding="utf-8") as f:
        json.dump(dicionario, f, ensure_ascii=False, indent=4)

def gerar_codigo_cutter(sobrenome, dicionario_cutter):
    sobrenome = sobrenome.upper()
    if sobrenome in dicionario_cutter:
        return dicionario_cutter[sobrenome]

    prefixo = sobrenome[:3]
    numero = 1
    codigo = f"{prefixo}{str(numero).zfill(3)}"

    codigos_existentes = set(dicionario_cutter.values())
    while codigo in codigos_existentes:
        numero += 1
        codigo = f"{prefixo}{str(numero).zfill(3)}"

    dicionario_cutter[sobrenome] = codigo
    salvar_dicionario_cutter(dicionario_cutter)
    return codigo

def buscar_dados_isbn(isbn: str):
    url = f"https://brasilapi.com.br/api/isbn/v1/{isbn}"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code != 200:
            return None
        return response.json()
    except Exception:
        return None

def cadastrar_exemplares(isbn: str, sigla: str, quantidade: int) -> pd.DataFrame:
    info = buscar_dados_isbn(isbn)
    if not info:
        raise ValueError("ISBN não encontrado na BrasilAPI.")

    titulo = info.get("title", "")
    autor = ", ".join(info.get("authors", [])) if info.get("authors") else ""
    ano_publicacao = info.get("year", "")
    cdd = info.get("classification", "")

    dicionario_cutter = carregar_dicionario_cutter()

    exemplares = []
    for i in range(1, quantidade + 1):
        sobrenome = autor.split(" ")[-1] if autor else "AUT"
        cutter = gerar_codigo_cutter(sobrenome, dicionario_cutter)
        exemplar = {
            "Título": titulo,
            "Autor": autor,
            "ISBN": isbn,
            "ID_Acervo": f"{sigla}{str(i).zfill(5)}",
            "Exemplar": f"Ex.{i}   {sigla}",
            "Código Cutter": cutter,
            "Ano de Publicação": ano_publicacao,
            "Classificação CDD": cdd
        }
        exemplares.append(exemplar)

    df = pd.DataFrame(exemplares)
    return df
