from flask import Flask, request, jsonify
import re
from cadastro import cadastrar_exemplares
from gerar_docx import gerar_documento
from sincronizar import sincronizar_acervo


app = Flask(__name__)

@app.route("/executar-comando", methods=["POST"])
def executar_comando():
    body = request.get_json()
    comando = body.get("comando", "") if body else ""

    if not comando:
        return jsonify({"erro": "Você precisa enviar um campo 'comando' no corpo da requisição."}), 400

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
        registros = cadastrar_exemplares(isbn, quantidade, sigla)
        gerar_documento(registros)
        sincronizar_dados(registros, sigla)
    except Exception as e:
        return jsonify({"erro": f"Erro ao processar o comando: {str(e)}"}), 500

    return jsonify({
        "status": "ok",
        "quantidade": len(registros),
        "dados": registros
    })

# Isso aqui é **ESSENCIAL** para o Render/Gunicorn conseguir importar `app`
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=3000)
