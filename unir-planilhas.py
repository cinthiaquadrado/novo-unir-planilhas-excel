import os
import pandas as pd

# Pasta onde estão as planilhas
pasta = '/caminho/para/a/pasta/'

# Nome da aba a ser lida em cada planilha
aba = 'Nome da Aba'

# Lista vazia para armazenar os DataFrames
dfs = []

# Loop pelas planilhas na pasta
for arquivo in os.listdir(pasta):
    if arquivo.endswith('.xlsx'):
        # Lê a planilha e a aba especificada em um DataFrame
        df = pd.read_excel(os.path.join(pasta, arquivo), sheet_name=aba)
        # Adiciona uma coluna com o nome do arquivo de origem
        df['Nome do Arquivo'] = arquivo
        # Adiciona o DataFrame à lista
        dfs.append(df)

# Concatena todos os DataFrames em um único DataFrame
resultado = pd.concat(dfs, ignore_index=True)

# Salva o resultado em um arquivo Excel
resultado.to_excel('arquivo_final.xlsx', index=False)
