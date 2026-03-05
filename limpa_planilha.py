import re
import pandas as pd
from tkinter import filedialog
import tkinter as tk


root = tk.Tk()
root.withdraw()

caminho_do_arquivo = filedialog.askopenfilename(
    title="Selecione a planilha para limpar",
    filetypes=[("Arquivos de Excel", "*.xlsx *.xls"), ("Arquivos CSV", "*.csv")]
)
def extrair_codigo(texto):
    if pd.isna(texto):
        return ""

    padrao_cnj = r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}'
    match = re.search(padrao_cnj, str(texto))

    if match:
        return match.group()

    return ""

if caminho_do_arquivo:
    print(f"Arquivo selecionado: {caminho_do_arquivo}")
    df = pd.read_excel(caminho_do_arquivo)
    df['Processo (padrão sistema vivo)'] = df['Processo (padrão sistema vivo)'].apply(extrair_codigo)
    df.drop_duplicates(subset=['Processo (padrão sistema vivo)'], keep='first', inplace=True)
    df.to_excel('planilha_base.xlsx', index=False)


else:
    print("Nenhum arquivo foi selecionado.")



