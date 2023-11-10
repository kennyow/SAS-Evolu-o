import pandas as pd

# Nome do arquivo Excel de entrada
arquivo_xlsx = r"C:\Users\citee\OneDrive\Documentos\VSCODE\Python\SAS_EVOLUCAO2.xlsx"

# Nome do arquivo CSV de saída
arquivo_csv = "TURMASSAS.csv"

# Carregando o arquivo Excel em um DataFrame do pandas
df = pd.read_excel(arquivo_xlsx)

# Salvando o DataFrame em um arquivo CSV
df.to_csv(arquivo_csv,  encoding='latin-1',index=False)  # O argumento index=False evita que os índices sejam incluídos no CSV

print(f"Conversão concluída. Dados salvos em '{arquivo_csv}'.")
