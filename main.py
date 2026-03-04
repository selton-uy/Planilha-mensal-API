import pandas as pd
from monta_tabela import pecorrer_tabela

df = pd.read_excel('planilha_base.xlsx', dtype=str)
COLUNA_CODIGO = 'Processo (padrão sistema vivo)'

df.dropna(subset=[COLUNA_CODIGO], inplace=True)
df.drop_duplicates(subset=[COLUNA_CODIGO], inplace=True)

pecorrer_tabela(df, COLUNA_CODIGO)