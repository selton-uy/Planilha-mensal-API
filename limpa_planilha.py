import re
import pandas as pd


PADRAO_CNJ = r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}'


def _formatar_cnj(digitos: str) -> str:
    """Recebe 20 dígitos e formata no padrão CNJ: NNNNNNN-DD.AAAA.J.TT.OOOO"""
    return f"{digitos[0:7]}-{digitos[7:9]}.{digitos[9:13]}.{digitos[13]}.{digitos[14:16]}.{digitos[16:20]}"


def _reconstruir_de_numero(texto: str) -> str:
    """
    Tenta recuperar um código CNJ que o Excel corrompeu para número
    (ex: notação científica 1.23457e+19 ou número sem zeros iniciais).
    O código CNJ tem sempre exatamente 20 dígitos.
    """
    try:
        # float() lida com notação científica; int() remove a parte decimal
        numero = int(float(texto))
        digitos = str(numero).zfill(20)  # recoloca zeros à esquerda
        if len(digitos) == 20:
            codigo = _formatar_cnj(digitos)
            # Valida que o resultado bate no padrão antes de retornar
            if re.fullmatch(PADRAO_CNJ, codigo):
                return codigo
    except (ValueError, TypeError):
        pass
    return ""


def extrair_codigo(texto) -> str:
    if pd.isna(texto):
        return ""

    texto_str = str(texto).strip()

    # Caso 1: já está no formato correto
    match = re.search(PADRAO_CNJ, texto_str)
    if match:
        return match.group()

    # Caso 2: Excel converteu para número (científico ou inteiro sem zeros)
    return _reconstruir_de_numero(texto_str)


def main(caminho_do_arquivo: str):
    """
    Recebe o caminho do arquivo Excel selecionado pela interface,
    corrige e limpa os códigos de processo e salva como planilha_base.xlsx.
    """
    df = pd.read_excel(caminho_do_arquivo)
    df['Processo (padrão sistema vivo)'] = df['Processo (padrão sistema vivo)'].apply(extrair_codigo)
    df.drop_duplicates(subset=['Processo (padrão sistema vivo)'], keep='first', inplace=True)
    df.to_excel('planilha_base.xlsx', index=False)
    print(f"planilha_base.xlsx salva com {len(df)} registros.")