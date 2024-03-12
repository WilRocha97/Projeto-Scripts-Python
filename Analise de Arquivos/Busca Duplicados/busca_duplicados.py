import os
import hashlib

def calcular_hash_arquivo(caminho_arquivo, tamanho_bloco=65536):
    sha = hashlib.sha256()
    with open(caminho_arquivo, 'rb') as f:
        bloco = f.read(tamanho_bloco)
        while len(bloco) > 0:
            sha.update(bloco)
            bloco = f.read(tamanho_bloco)
    return sha.hexdigest()

def encontrar_duplicatas(pasta):
    hash_map = {}
    duplicatas = []

    for pasta_raiz, _, arquivos in os.walk(pasta):
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta_raiz, arquivo)
            hash_arquivo = calcular_hash_arquivo(caminho_arquivo)

            if hash_arquivo in hash_map:
                duplicatas.append((caminho_arquivo, hash_map[hash_arquivo]))
            else:
                hash_map[hash_arquivo] = caminho_arquivo

    return duplicatas

pasta_alvo = 'C:\\Users\\willianr\\Documents\\Consulta Certidão Negativa Receita Federal'
duplicatas_encontradas = encontrar_duplicatas(pasta_alvo)

if duplicatas_encontradas:
    print("Duplicatas encontradas:")
    for duplicata in duplicatas_encontradas:
        print(f"Arquivo 1: {duplicata[0]}")
        print(f"Arquivo 2: {duplicata[1]}")
        print("-----------")
else:
    print("Nenhuma duplicata encontrada.")
