import pandas as pd
import re
import unicodedata
import json
import requests
import shutil

# Carrega dicionário Cutter externo
with open("cutter_dict.json", encoding="utf-8") as f:
    cutter_dict = json.load(f)

ARTIGOS = {"o", "os", "a", "as", "um", "uns", "uma", "umas", "do", "da", "dos", "das", "de", "em", "no", "na", "nos", "nas"}

def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def extrair_sobrenome(autor):
    partes = autor.split()
    return remover_acentos(partes[-1]) if partes else "Anon"

def extrair_letras_uteis_titulo(titulo):
    palavras = re.findall(r"\b\w+\b", remover_acentos(titulo.lower()))
    uteis = [p for p in palavras if p not in ARTIGOS]
    return uteis if uteis else ["x"]

def buscar_codigo_cutter(sobrenome):
    sobrenome_norm = sobrenome.lower()
    for i in range(6, 1, -1):
        chave = sobrenome_norm[:i]
        if chave in cutter_dict:
            return cutter_dict[chave]
    return "000"

def gerar_codigo_cutter(df, autor, titulo):
    sobrenome = extrair_sobrenome(autor)
    codigo = buscar_codigo_cutter(sobrenome)
    letras_titulo = extrair_letras_uteis_titulo(titulo)

    sufixo = letras_titulo[0][0].lower()
    indice = 1
    while any((str(codigo) + sufixo) == str(e)[-4:].lower() for e in df.get("Código Cutter", [])):
        if indice >= len(letras_titulo[0]):
            if len(letras_titulo) > 1:
                letras_titulo.pop(0)
                indice = 0
            else:
                sufixo += "x"
                break
        sufixo = letras_titulo[0][indice].lower()
        indice += 1

    return f"{sobrenome[0].upper()}{codigo}{sufixo}"

def enriquecer_dados(isbn: str) -> dict:
    url = f"https://brasilapi.com.br/api/isbn/v1/{isbn}"
    try:
        r = requests.get(url, timeout=5)
        if r.status_code != 200:
            return {}

        data = r.json()

        return {
            "Título": data.get("title", "").strip(),
            "Subtítulo": data.get("subtitle", "").strip(),
            "Autor": ", ".join(data.get("authors") or []).strip(),
            "Editora": data.get("publisher", "").strip(),
            "Ano de Publicação": str(data.get("year", "")),
            "Local de Publicação": data.get("location", "").strip(),
            "Edição": data.get("edition", "").strip(),
            "Volume": data.get("volume", "").strip(),
            "Classificação CDD": data.get("classification", "").strip(),
            "ISBN": isbn
        }

    except Exception as e:
        print(f"⚠️ Erro ao consultar BrasilAPI: {e}")
        return {}

def cadastrar_varios_exemplares_por_comando(comando_usuario: str) -> pd.DataFrame:
    match = re.search(r"cadastrar\s+(\d+)\s+exemplares.*isbn[:\s]*([\d\-]+).*sigla[:\s]*([A-Z]{2,4})", comando_usuario, re.IGNORECASE)
    if not match:
        raise ValueError("❌ Comando inválido. Use: 'Cadastrar X exemplares da obra com ISBN: Y e sigla: Z'")

    quantidade = int(match.group(1))
    isbn = match.group(2).replace("-", "")
    sigla = match.group(3).upper()

    dados_base = enriquecer_dados(isbn)
    if not dados_base:
        raise Exception("❌ Não foi possível obter os dados da obra via ISBN.")

    registros = []
    df_existente = pd.DataFrame()
    for i in range(quantidade):
        n = i + 1
        registro = dados_base.copy()
        registro["ID_Acervo"] = f"{sigla}{str(n).zfill(5)}"
        registro["Código Cutter"] = gerar_codigo_cutter(df_existente, registro["Autor"], registro["Título"])
        registro["Exemplar"] = f"Ex.{n}   {sigla}"
        registros.append(registro)
        df_existente = pd.concat([df_existente, pd.DataFrame([registro])], ignore_index=True)

    df_final = pd.DataFrame(registros)
    df_final.to_excel("acervo.xlsx", index=False)

    shutil.move("acervo.xlsx", "/mnt/data/acervo.xlsx")

    return df_final