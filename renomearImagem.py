import os

pasta = r"C:\Users\citee\OneDrive\Documentos\Scanned Documents"
novo_nome_base = "3B_sas"

def renomear_arquivos(pasta, novo_nome_base):
    # Verifica se a pasta existe
    if not os.path.exists(pasta):
        print("A pasta especificada não existe.")
        return

    # Percorre todos os arquivos na pasta
    for nome_arquivo in os.listdir(pasta):
        if nome_arquivo.startswith("Imagem") and nome_arquivo.lower().endswith(".jpg"):
            # Constrói o novo nome para o arquivo
            novo_nome = f"{novo_nome_base}_{nome_arquivo[len('Imagem'):]}"
            
            # Caminho completo dos arquivos de entrada e saída (salvando na mesma pasta)
            caminho_antigo = os.path.join(pasta, nome_arquivo)
            caminho_novo = os.path.join(pasta, novo_nome)

            # Renomeia o arquivo
            os.rename(caminho_antigo, caminho_novo)
            print(f"Arquivo renomeado: {caminho_antigo} -> {caminho_novo}")

# Chama a função para renomear os arquivos na pasta especificada
renomear_arquivos(pasta, novo_nome_base)
