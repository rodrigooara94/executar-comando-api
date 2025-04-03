import pandas as pd
import re

def extrair_numeracao_id(id_str):
    """Extrai prefixo e número do ID_Acervo (ex: BEN00012 → BEN, 12)"""
    match = re.match(r"([A-Z]{3})(\d+)", str(id_str))
    if match:
        return match.group(1), int(match.group(2))
    return None, None

def sincronizar_acervo(df, coluna_id='ID_Acervo'):
    """
    Apenas analisa o acervo. Não altera o DataFrame.
    
    Retorna:
        - prefixo (assumido como o primeiro válido)
        - último ID numérico encontrado
        - próximo ID sugerido
        - lacunas na sequência (se houver)
        - relatório (DataFrame com os IDs faltando)
    """
    if coluna_id not in df.columns or df.empty:
        return {
            'prefixo': None,
            'ultimo_id': 0,
            'faltando': [],
            'proximo_id_sugerido': None,
            'relatorio': pd.DataFrame()
        }

    extraidos = df[coluna_id].dropna().map(extrair_numeracao_id)
    prefixos, numeros = zip(*extraidos)

    # Assume o primeiro prefixo válido como base
    prefixo_padrao = next((p for p in prefixos if p), None)
    numeros_validos = sorted(set(filter(None, numeros)))

    if not numeros_validos:
        return {
            'prefixo': prefixo_padrao,
            'ultimo_id': 0,
            'faltando': [],
            'proximo_id_sugerido': f"{prefixo_padrao}00001" if prefixo_padrao else None,
            'relatorio': pd.DataFrame()
        }

    ultimo_id = max(numeros_validos)
    faltando = sorted(set(range(1, ultimo_id + 1)) - set(numeros_validos))
    proximo_id = f"{prefixo_padrao}{str(ultimo_id + 1).zfill(5)}" if prefixo_padrao else None

    return {
        'prefixo': prefixo_padrao,
        'ultimo_id': ultimo_id,
        'faltando': faltando,
        'proximo_id_sugerido': proximo_id,
        'relatorio': pd.DataFrame({'IDs faltando': faltando}) if faltando else pd.DataFrame()
    }

def carregar_todo_excel(caminho_excel):
    """Carrega todo o conteúdo de um arquivo Excel"""
    return pd.read_excel(caminho_excel, dtype=str)

def salvar_em_excel(df: pd.DataFrame, nome_arquivo: str):
    """Salva um DataFrame em um arquivo Excel"""
    df.to_excel(nome_arquivo, index=False)
