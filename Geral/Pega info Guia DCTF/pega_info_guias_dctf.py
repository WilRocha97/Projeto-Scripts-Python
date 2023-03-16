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
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        
        with fitz.open(arq) as pdf:

            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                if count == 1:
                    continue
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # print(textinho)
                    # Procura o valor a recolher da empresa
                    try:
                        cnpj = re.compile(r'Documento de Arrecadação\nde Receitas Federais\n(\d.+)\n').search(textinho).group(1)
                    except:
                        cnpj = re.compile(r'Documento de Arrecadação\ndo eSocial\n(\d.+)\n').search(textinho).group(1)
                        
                    # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a essa info
                    try:
                        valor = re.compile(r'Valor Total do Documento\n(.+)\nCNPJ').search(textinho).group(1)
                    except:
                        valor = re.compile(r'Valor Total do Documento\n(.+)\nCPF').search(textinho).group(1)
                        
                    print(f'{cnpj} - {valor}')

                    _escreve_relatorio_csv(f"{cnpj};{valor}")

                except():
                    print(textinho)
                    

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv(';'.join(['CNPJ', 'VALOR']))
    print(datetime.now() - inicio)
