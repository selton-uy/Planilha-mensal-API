from busca_dados import *


def pecorrer_tabela(tabela, coluna_codigo):
    for index, row in tabela.iterrows():
        try:
            codigo = str(row[coluna_codigo])
            if not codigo or codigo.lower() == 'nan':
                continue

            processo = consultar_processo(row[coluna_codigo])
            dados_processo = extrair_dados(processo)

            recorrente = str(dados_processo.get('RECORRENTE') or "").upper()

            tabela.at[index, 'Parte autora'] = dados_processo.get('RECORRENTE')
            tabela.at[index, 'Turma/Câmara'] = dados_processo.get('TURMA')
            tabela.at[index, 'Relator - Magistrado - 2ª Instância'] = dados_processo.get('RELATOR')
            tabela.at[index, 'Fundamentação principal'] = dados_processo.get('SUMULA')

            if 'TELEFONICA' in recorrente and 'E OUTRO' in recorrente:
                resultado = 'AMBOS'
            elif 'TELEFONICA' in recorrente:
                resultado = 'VIVO'
            elif recorrente == '':
                resultado = ''
            else:
                resultado = 'PARTE AUTORA'

            tabela.at[index, 'Recorrente (Vivo, Parte Autora, Ambos)'] = resultado

            print(f'Sucesso no processo: {codigo}')

        except Exception as e:
            print(f'Erro ao processar índice {index} (Código: {row[coluna_codigo]}): {e}')
            continue
    tabela.to_excel('planilha_nova.xlsx', index=False)