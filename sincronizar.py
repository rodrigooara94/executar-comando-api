import re
import requests
from gerar_docx import gerar_documento

def sincronizar_dados(comando):
    try:
        # Expressão regular mais flexível e case-insensitive
        pattern = re.compile(
            r'cadastrar\s+(\d+)\s+exemplares?\s+(?:da\s+obra\s+)?com\s+isbn\s*:\s*([\d\-Xx]+)\s+e\s+sigla\s*:\s*([A-Za-z]+)',
            re.IGNORECASE
        )
        match = pattern.search(comando)
        if not match:
            return {"erro": "Comando inválido. Use: Cadastrar X exemplares da obra com ISBN: Y e sigla: Z"}

        quantidade = int(match.group(1))
        isbn = match.group(2).replace('-', '')
        sigla = match.group(3).upper()

        # Requisição à BrasilAPI
        response = requests.get(f"https://brasilapi.com.br/api/isbn/v1/{isbn}")
        if response.status_code != 200:
            return {"erro": "ISBN não encontrado na BrasilAPI."}

        dados_livro = response.json()
        titulo = dados_livro.get("title", "Título desconhecido")
        autor = ', '.join(dados_livro.get("authors", [])) or "Autor desconhecido"

        gerar_documento(quantidade, titulo, autor, sigla)
        return {"mensagem": f"{quantidade} exemplares da obra '{titulo}' cadastrados com sucesso com a sigla '{sigla}'."}

    except Exception as e:
        return {"erro": f"Erro ao processar o comando: {str(e)}"}
