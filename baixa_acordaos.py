import os
from html import unescape
from datetime import datetime

import pandas as pd
import requests

from busca_dados import consultar_processo


def _extrair_ged(json_processo: dict) -> str | None:
    inteiro_teor = json_processo.get('InteiroTeor', [])
    candidatos = [
        item for item in inteiro_teor
        if item.get('Descr', '').strip() == 'Súmula de Julgamento'
    ]
    if not candidatos:
        return None
    if len(candidatos) == 1:
        return candidatos[0].get('ArqGED')

    def parse_data(item):
        try:
            return datetime.strptime(item.get('DtHrMovStr', ''), '%d/%m/%Y')
        except ValueError:
            return datetime.min

    return max(candidatos, key=parse_data).get('ArqGED')


def _baixar_pdf(ged: str, caminho_arquivo: str):
    url = unescape(
        f"https://www3.tjrj.jus.br/gedcacheweb/default.aspx?UZIP=1&amp;GEDID={ged}&amp;USER="
    )
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www3.tjrj.jus.br/",
    }
    response = requests.get(url, stream=True, headers=headers)
    response.raise_for_status()
    with open(caminho_arquivo, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)


def main(caminho_planilha: str, coluna_codigo: str, pasta_destino: str, log=print):
    df = pd.read_excel(caminho_planilha)
    df = df.astype(str).replace('nan', '')

    total = len(df)
    indices_falha = []
    os.makedirs(pasta_destino, exist_ok=True)

    for index, row in df.iterrows():
        num = str(int(index) + 1).zfill(2)
        codigo = row.get(coluna_codigo, '').strip()

        if not codigo:
            continue

        try:
            processo = consultar_processo(codigo)
            ged = _extrair_ged(processo)

            if not ged:
                log(f"⚠️  [{num}/{total}] Sem 'Súmula de Julgamento': {codigo}")
                indices_falha.append(index)
                continue

            nome_arquivo = os.path.join(pasta_destino, f"{codigo.replace('/', '_')}.pdf")
            _baixar_pdf(ged, nome_arquivo)
            log(f"✅ [{num}/{total}] Baixado: {codigo}")

        except Exception as e:
            log(f"❌ [{num}/{total}] Erro em {codigo}: {e}")
            indices_falha.append(index)
            continue

    # Salva planilha só com as falhas
    if indices_falha:
        df_falhas = df.loc[indices_falha]
        caminho_falhas = os.path.join(pasta_destino, "acordaos_falhos.xlsx")
        df_falhas.to_excel(caminho_falhas, index=False)
        log(f"📋 Planilha de falhas salva em: {caminho_falhas}")

    log(f"💾 Downloads concluídos em: {pasta_destino}")
    if indices_falha:
        log(f"⚠️  {len(indices_falha)} processo(s) não foram baixados.")
    else:
        log("🎉 Todos os acórdãos baixados com sucesso!")

    return len(indices_falha)