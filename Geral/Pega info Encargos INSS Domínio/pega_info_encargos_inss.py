# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

def analiza():
    arq = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    # Abrir o pdf
    with fitz.open(arq) as pdf:

        # Para cada página do pdf, se for a segunda página o script ignora
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # print(textinho)
                # time.sleep(55)
                # Procura o valor a recolher da empresa
                cnpj = re.compile(r'Empresa:\n(.+)').search(textinho).group(1)

                # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a essa info
                try:
                    valor = re.compile(r'(.+\d)\nTotal\nSistema').search(textinho).group(1)
                except:
                    continue
                    
                print(f'{cnpj} - {valor}')

                _escreve_relatorio_csv(f"{cnpj};{valor}", nome='Encargos INSS Domínio web')

            except():
                print(textinho)
                print('ERRO')
        

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ;VALOR DOMÍNIO', nome='Encargos INSS Domínio web')
    print(datetime.now() - inicio)
