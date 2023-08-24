# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

def analiza():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arquivo in os.listdir(documentos):
        print(f'\nArquivo: {arquivo}')
        # Abrir o pdf
        arq = os.path.join(documentos, arquivo)
        
        with fitz.open(arq) as pdf:
            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                if count == 1:
                    continue
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    
                    # print(textinho)
                    # time.sleep(44)
                    
                    dados = re.compile(r'(.+)\n(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)\nCNPJ').search(textinho)
                    if not dados:
                        dados = re.compile(r'(.+)\n(\d+)\nCEI').search(textinho)
                        
                    nome = dados.group(1)
                    cnpj = dados.group(2)
                    
                    funcionario = re.compile(r'(.+)\nNome do Funcionário').search(textinho).group(1)
                    valor = re.compile(r'(.+,\d+)\nCódigo').search(textinho).group(1)
                    
                    _escreve_relatorio_csv(f'{cnpj};{nome};{funcionario};{valor}', nome='Analise Pro-Labore')
                    
                except:
                    print(textinho)


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ;EMPRESA;NOME;VALOR', nome='Analise Pro-Labore')
    print(datetime.now() - inicio)
