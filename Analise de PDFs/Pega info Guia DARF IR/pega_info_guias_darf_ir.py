# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

# Definir os padrões de regex

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
                    
                    # if arquivo == 'RAFAEL BUCHALLA BAGARELLI FERREIRA - 22031924885 - 0561 - DARF CUCA.pdf':
                        # print(textinho)
                    
                    # Procura infos da empresa
                    try:
                        # Domínio CNPJ Antigo
                        infos = re.compile(r'(\d\d/\d\d/\d\d\d\d)\n(.+)\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)').search(textinho)
                        cnpj = infos.group(3)
                        nome = infos.group(2)
                        apuracao = infos.group(1)
                    except:
                        try:
                            #Domínio CPF Antigo
                            infos = re.compile(r'(\d\d/\d\d/\d\d\d\d)\n(.+)\n(\d\d\d.\d\d\d.\d\d\d-\d\d)').search(textinho)
                            cnpj = infos.group(3)
                            nome = infos.group(2)
                            apuracao = infos.group(1)
                        except:
                            try:
                                # DPCUCA antigo CNPJ
                                infos = re.compile(r'Período de Apuração.+\n(\d\d/\d\d/\d\d\d\d)\n.+\n.+\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)').search(textinho)
                                apuracao = infos.group(1)
                                cnpj = infos.group(2)
                                nome = re.compile(r'\d\d/\d\d/\d\d\d\d\n(.+)\n +').search(textinho).group(1)
                            except:
                                try:
                                    # DPCUCA antigo CPF
                                    infos = re.compile(r'Período de Apuração.+\n(\d\d/\d\d/\d\d\d\d)\n.+\n.+\n(\d\d\d.\d\d\d.\d\d\d-\d\d)').search(textinho)
                                    apuracao = infos.group(1)
                                    cnpj = infos.group(2)
                                    nome = re.compile(r'\d\d/\d\d/\d\d\d\d\n(.+)\n +').search(textinho).group(1)
                                except:
                                    # Domínio Novo
                                    apuracao = re.compile(r'Razão Social\n(\d\d/\d\d/\d\d\d\d)').search(textinho).group(1)
                                    try:
                                        infos = re.compile(r'Documento de Arrecadação\nde Receitas Federais\n(\d.+)\n(.+)').search(textinho)
                                        cnpj = infos.group(1)
                                        nome = infos.group(2)
                                    except:
                                        infos = re.compile(r'Documento de Arrecadação\ndo eSocial\n(\d.+)\n(.+)').search(textinho)
                                        cnpj = infos.group(1)
                                        nome = infos.group(2)
                                    
                    # Procura valor
                    try:
                        valor = re.compile(r'SECRETARIA DA RECEITA FEDERAL DO BRASIL\n(.+,\d+)').search(textinho).group(1)
                    except:
                        try:
                            valor = re.compile(r'Valor Total\n +(.+,\d+)\n').search(textinho).group(1)
                        except:
                            try:
                                valor = re.compile(r'Valor Total do Documento\n(.+)\nCNPJ').search(textinho).group(1)
                            except:
                                valor = re.compile(r'Valor Total do Documento\n(.+)\nCPF').search(textinho).group(1)
                        
                    # Procura código da receita
                    try:
                        codigo = re.compile(r'\n(\d\d\d\d)\n').search(textinho).group(1)
                    except:
                        codigo = re.compile(r'Receitas Federais\n(.+)\n').search(textinho).group(1)
                        codigo = codigo.replace(' ','')
                    
                    print(f'{valor} - {apuracao} - {codigo} - {cnpj} - {nome}')
                    _escreve_relatorio_csv(f"{cnpj};{nome};{apuracao};{codigo};{valor}", nome='Info Guias DARF IR')

                except():
                    print(textinho)
                    

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ;NOME;APURAÇÃO;CÓDIGO;VALOR', nome='Info Guias DARF IR')
    print(datetime.now() - inicio)
