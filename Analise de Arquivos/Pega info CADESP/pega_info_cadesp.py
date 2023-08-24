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
        _escreve_relatorio_csv(f"{arq}")
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:

            # Para cada página do pdf
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # print(textinho)
                try:
                    infos = re.findall(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\n\d\d)\n(\d\d\d\.\d\d\d\.\d\d\d\.\d\d\d)\n(.+)\n(.+)', textinho)
                    for info in infos:
                        cnpj = info[0].replace('\n', '')
                        ie = info[1]
                        nome = info[2]
                        situacao = info[3]
                        if situacao != 'Ativo' and situacao != 'Não Ativo':
                            situacao = re.search(r'(' + info[0] + ')\n(' + info[1] + ')\n(' + info[2] + ')\n(' + info[3] + ')\n(.+)', textinho).group(5)
                        
                        _escreve_relatorio_csv(f"{cnpj};{ie};{nome};{situacao}")
                except:
                    print('ERRO')

                try:
                    infos = re.findall(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-)(\d\d\d\.\d\d\d\.\d\d\d\.\d\d\d)\n(.+)\n(.+)', textinho)
                    for info in infos:
                        cnpj = info[0].replace('\n', '')
                        ie = info[1]
                        nome = info[2]
                        situacao = info[3]
                        if situacao != 'Ativo' and situacao != 'Não Ativo':
                            situacao = re.search(r'(' + info[0] + ')(' + info[1] + ')\n(' + info[2] + ')\n(' + info[3] + ')\n(.+)', textinho).group(5)

                        _escreve_relatorio_csv(f"{cnpj};{ie};{nome};{situacao}")
                except:
                    print('ERRO')

                try:
                    infos = re.findall(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\n\d\d)\n(\d\d\d\.\d\d\d\.\d\d\d\.\d\d\d)(.+)\n(.+)', textinho)
                    for info in infos:
                        cnpj = info[0].replace('\n', '')
                        ie = info[1]
                        nome = info[2]
                        situacao = info[3]
                        if situacao != 'Ativo' and situacao != 'Não Ativo':
                            situacao = re.search(r'(' + info[0] + ')\n(' + info[1] + ')(' + info[2] + ')\n(' + info[3] + ')\n(.+)', textinho).group(5)

                        _escreve_relatorio_csv(f"{cnpj};{ie};{nome};{situacao}")
                except:
                    print('ERRO')

                try:
                    infos = re.findall(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-)(\d\d\d\.\d\d\d\.\d\d\d\.\d\d\d)(.+)\n(.+)', textinho)
                    for info in infos:
                        cnpj = info[0].replace('\n', '')
                        ie = info[1]
                        nome = info[2]
                        situacao = info[3]
                        if situacao != 'Ativo' and situacao != 'Não Ativo':
                            situacao = re.search(r'(' + info[0] + ')(' + info[1] + ')(' + info[2] + ')\n(' + info[3] + ')\n(.+)', textinho).group(5)

                        _escreve_relatorio_csv(f"{cnpj};{ie};{nome};{situacao}")
                except:
                    print('ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    pega_info()
    print(datetime.now() - inicio)
