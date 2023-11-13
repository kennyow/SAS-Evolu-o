import os

pasta = r"C:\Users\citee\OneDrive\Documentos\Scanned Documents\AAS -6 ANOS -  28 -08- 23"
novo_nome_base = "Imagem "

def renomear_arquivos_jpg(pasta, novo_nome_base):
    # Verifica se a pasta existe
    if not os.path.exists(pasta):
        print("A pasta especificada não existe.")
        return

    # Percorre todos os arquivos na pasta
    for contador, nome_arquivo in enumerate(os.listdir(pasta)):
        if nome_arquivo.lower().endswith(".jpg"):
            # Constrói o novo nome para o arquivo
            novo_nome = f"{novo_nome_base}_{contador + 200}.jpg"

            # Caminho completo dos arquivos de entrada e saída (salvando na mesma pasta)
            caminho_antigo = os.path.join(pasta, nome_arquivo)
            caminho_novo = os.path.join(pasta, novo_nome)

            # Renomeia o arquivo
            os.rename(caminho_antigo, caminho_novo)
            print(f"Arquivo renomeado: {caminho_antigo} -> {caminho_novo}")

# Chama a função para renomear os arquivos .jpg na pasta especificada
renomear_arquivos_jpg(pasta, novo_nome_base)
