from flask import Flask, request, jsonify
import re
import pandas as pd
from cadastro import cadastrar_exemplares
from gerar_docx import gerar_documento
from sincronizar import sincronizar_dados

app = Flask(__name__)

@app.route("/")
def home():
    return "API do Executar Comando está online!"

@app.route("/executar-comando", methods=["POST"])
def executar_comando():
    body = request.get_json()
    comando = body.get("comando", "") if body else ""

    if not comando:
        return jsonify({"erro": "Você precisa enviar um campo 'comando' no corpo da requisição."}), 400

    # Extrair dados do comando em linguagem natural
    match = re.search(
        r"cadastrar\s+(\d+)\s+exemplares.*isbn[:\s]*([\d\-]+).*sigla[:\s]*([A-Z]{2,4})",
        comando,
        re.IGNORECASE
    )

    if not match:
        return jsonify({
            "erro": "Comando inválido. Use: Cadastrar X exemplares da obra com ISBN: Y e sigla: Z"
        }), 400

    quantidade = int(match.group(1))
    isbn = match.group(2).replace("-", "")
    sigla = match.group(3).upper()

    try:
        # Gerar exemplares
        df = cadastrar_exemplares(isbn, sigla, quantidade)
        registros = df.to_dict(orient="records")

        # Gerar documento .docx
        nome_docx = f"{sigla}_{isbn}.docx"
        gerar_documento(df, nome_docx)

        # Gerar planilha Excel com nome fixo
        nome_excel = f"acervo_{sigla}.xlsx"
        sincronizar_dados(df, nome_arquivo=nome_excel)

    except Exception as e:
        return jsonify({"erro": f"Erro durante o processamento: {str(e)}"}), 500

    return jsonify({
        "status": "ok",
        "quantidade": len(registros),
        "arquivo_docx": nome_docx,
        "arquivo_excel": nome_excel,
        "dados": registros
    })
