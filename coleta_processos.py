from busca_dados import *
from monta_tabela import criar_tabela
import pandas as pd
import os

path_entrada = 'planilha_vazia.xlsx'
path_saida = 'planilha_base.xlsx'
coluna_codigo = 'Processo (padrão sistema vivo)'
caluna_data = 'Data Publicação'

#nomes = ['TELEFONICA BRASIL S.A', 'vivo']
#todo pesquisar pelos dois nomes da vivo
criar_tabela()

df = pd.read_excel(path_entrada)

lista_de_processos = consultar_processo_por_nome_e_data('VIVO', 2026, 2026)
for index, processo in enumerate(lista_de_processos):
    print(f"Processando o prcesso: {index + 1} de {len(lista_de_processos)+1}")
    try:
        codigo = processo['codCnj']
        dados = consultar_processo_por_numero(codigo)

        movimentacoes = dados.get("DadosUltMovimentos", [])
        data_publicacao = None

        for item in movimentacoes:
            if item.get("Descr") == "Data de Publicação":
                data_publicacao = item.get("Valor")
                break

        if data_publicacao:
            df.at[index, coluna_codigo] = codigo
            df.at[index, caluna_data] = data_publicacao
        else:
            df.at[index, coluna_codigo] = codigo
            df.at[index, caluna_data] = ''

            print(f"Aviso: Data não encontrada para o processo {codigo}")

    except Exception as e:
        print(f"Erro no índice {index}: {e}")

df.to_excel(path_saida, index=False)

if os.path.exists(path_entrada):
    os.remove(path_entrada)
else:
    print('NOT FOUND')