import os, time
import fitz
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir

documentos = ask_for_dir()
pasta_final = os.path.join(documentos, 'Editados')
os.makedirs(pasta_final, exist_ok=True)


def concatena(variavel, quantidade, posicao, caractere):
    variavel = str(variavel)
    if posicao == 'depois':
        # concatena depois
        while len(str(variavel)) < quantidade:
            variavel += str(caractere)
    if posicao == 'antes':
        # concatena antes
        while len(str(variavel)) < quantidade:
            variavel = str(caractere) + str(variavel)
    
    return variavel


for file in os.listdir(documentos):
    contador = 0
    conteudo_modificado = ''
    if not file.endswith('.RFB'):
        continue
    arq = os.path.join(documentos, file)
    with open(arq, 'r') as arquivo:
        # Leia o conteúdo do arquivo
        conteudo = arquivo.readlines()

        # verifica os números de chave que podem já existir no arquivo original
        lista_chaves = []
        for count, linha in enumerate(conteudo[:-1], start=1):
            codigo_capturado = linha[34:38]
            if str(codigo_capturado) == '5952' or str(codigo_capturado) == '1708' or str(codigo_capturado) == '3208' or str(codigo_capturado) == '5706' or str(codigo_capturado) == '8045' or str(codigo_capturado) == '2991':
                contador += 1
                continue
            else:
                conteudo_modificado += linha
            
            linhas = concatena(str(count + 1 - contador), 4, 'antes', '0')
        
        print(linhas, count + 1, contador)
        ultima_linha = conteudo[-1].replace('\n', '').strip()
        print(ultima_linha)
        ultima_linha_limpa = ultima_linha[:-4]
        print(ultima_linha_limpa)
        ultima_linha_editada = ultima_linha_limpa + linhas + '\n'
        print(ultima_linha_editada)
    # Abra o arquivo para escrita
    with open(os.path.join(pasta_final, file), 'w') as arquivo:
        # Escreva o conteúdo modificado de volta no arquivo
        arquivo.writelines(conteudo_modificado + ultima_linha_editada)
    
    