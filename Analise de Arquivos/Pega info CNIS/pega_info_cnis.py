# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def pega_info():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:

            # Para cada página do pdf
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # print(textinho)
                # time.sleep(22)
                try:
                    contribuicoes = re.compile('(\d\d/\d\d\d\d)\n(\d\d/\d\d/\d\d\d\d)\n(\d.+,\d+)').findall(textinho)
                    for contribuicao in contribuicoes:
                        _escreve_relatorio_csv(f'{contribuicao[0]};{contribuicao[1]};{contribuicao[2]}', nome='Contribuições CNIS')
                except:
                    print('ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    pega_info()
    print(datetime.now() - inicio)
