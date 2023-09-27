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
            origen_anterior = ''
            # Para cada página do pdf
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                try:
                    contribuicoes = re.compile('(\d\d/\d\d\d\d)\n(\d\d/\d\d/\d\d\d\d)\n(.+,\d+)').findall(textinho)
                    for contribuicao in contribuicoes:
                        for i in range(200):
                            try:
                                origem = re.compile('Origem do Vínculo\n(.+\n){12}(.+\n){' + str(i) + '}(' + str(contribuicao[0]) + ')\n(' + str(contribuicao[1]) + ')\n(' + str(contribuicao[2]) + ')').search(textinho).group(1)
                                break
                            except:
                                origem = origen_anterior
                                
                        origem = origem.replace("\n", "")
                        _escreve_relatorio_csv(f'{contribuicao[0]};{contribuicao[1]};{contribuicao[2]};{origem}', nome='Contribuições CNIS')
                except:
                    print('ERRO')
                
                origens = re.compile(r'Origem do Vínculo\n(.+\n){12}').findall(textinho)
                for origen_nome in origens:
                    origen_anterior = origen_nome
    
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:
            origen_anterior = ''
            # Para cada página do pdf
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                try:
                    contribuicoes = re.compile('(\d\d/\d\d\d\d)\n(.+-.+)\n(.+-.+)\n(.+)\n(.+,\d+)').findall(textinho)
                    for contribuicao in contribuicoes:
                        for i in range(200):
                            try:
                                origem = re.compile('Origem do Vínculo\n(.+\n){11}(.+\n){' + str(i) + '}(' + str(contribuicao[0]) + ')\n(' + str(contribuicao[1]) + ')\n(' + str(contribuicao[2]) + ')\n(' + str(contribuicao[3]) + ')\n(' + str(contribuicao[4]) + ')').search(textinho).group(1)
                                break
                            except:
                                origem = origen_anterior
                        origem = origem.replace("\n", "")
                        _escreve_relatorio_csv(f'{contribuicao[0]};{contribuicao[1]};{contribuicao[2]};{contribuicao[3]};{contribuicao[4]};{origem}', nome='Contribuições CNIS Remunerações')
                except:
                    print('ERRO')
                
                origens = re.compile(r'Origem do Vínculo\n(.+\n){11}').findall(textinho)
                for origen_nome in origens:
                    origen_anterior = origen_nome
    
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:
            origen_anterior = ''
            # Para cada página do pdf
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                try:
                    remuneracoes = re.compile('\n(\d\d/\d\d\d\d)\n(.+,\d+)').findall(textinho)
                    for remuneracao in remuneracoes:
                        for i in range(200):
                            try:
                                origem = re.compile('Origem do Vínculo(\n.+){12}(\n.+){' + str(i) + '}\n(' + str(remuneracao[0]) + ')\n(' + str(remuneracao[1]) + ')').search(textinho).group(1)
                                break
                            except:
                                origem = origen_anterior
                        origem = origem.replace("\n", "")
                        _escreve_relatorio_csv(f'{remuneracao[0]};{remuneracao[1]};{origem}', nome='Contribuições CNIS Remunerações 2')
                except:
                    print('ERRO')
                
                origens = re.compile(r'Origem do Vínculo(\n.+){12}').findall(textinho)
                for origen_nome in origens:
                    origen_anterior = origen_nome
                    
                    
if __name__ == '__main__':
    inicio = datetime.now()
    pega_info()
    print(datetime.now() - inicio)
