from flask import Flask, request, jsonify
from cadastro import cadastrar_exemplares, carregar_todo_excel, salvar_em_excel

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return "API está ativa. Use POST em /executar para cadastrar exemplares."

@app.route("/executar", methods=["POST"])
def executar():
    try:
        data = request.get_json()
        comando = data.get("comando", "").lower()

        # Extrair valores usando regex com mais flexibilidade
        import re
        padrao = r"(?:cadastrar|adicionar|inserir)\s+(\d+)\s+(?:exemplar(?:es)?)\s+(?:da\s+obra\s+)?(?:com\s+)?isbn[:\s]*([\d\-xX]+)[\s,;]*.*sigla\s*(?:da biblioteca)?[:\s]*([a-zA-Z]+)"
        match = re.search(padrao, comando, re.IGNORECASE)

        if not match:
            return jsonify({
                "erro": "Comando inválido. Tente algo como: 'Cadastrar 3 exemplares da obra com ISBN: 1234567890 e sigla da biblioteca: ABC'"
            }), 400

        quantidade = int(match.group(1))
        isbn = match.group(2).replace("-", "")
        sigla = match.group(3).upper()

        # Cadastrar exemplares
        novos = cadastrar_exemplares(isbn, sigla, quantidade)

        # Carregar e salvar no Excel
        atual = carregar_todo_excel()
        combinado = pd.concat([atual, novos], ignore_index=True)
        salvar_em_excel(combinado)

        return jsonify({
            "mensagem": f"{quantidade} exemplar(es) cadastrados com sucesso para o ISBN {isbn} na biblioteca {sigla}."
        })

    except Exception as e:
        return jsonify({"erro": f"Erro ao processar o comando: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(debug=True)
