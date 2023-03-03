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
    for arq in os.listdir(documentos):
        # Abrir o pdf
        arquivo = os.path.join(documentos, arq)
        with fitz.open(arquivo) as pdf:
    
            pattern = re.compile(r' - \d+ - Analítico.+.pdf')
            empresa = re.sub(pattern, '', arq)

            cnpj = re.compile(r' - (\d+) - Analítico.+.pdf').search(arq).group(1)

            # Para cada página do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # Procura o valor a recolher da empresa
                    valores = None
                    indice = 20
                    while not valores:
                        indice = str(indice)
                        valores = re.compile(r'\n(.+)\n.+\n.0+(.+)\n(.+)\n(.+)\nSérie:\n(.+\n){' + indice + '}ISS\n.+\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)(\n.+){7}\nR\$ (.+)').findall(textinho)
                        indice = int(indice)
                        indice += 1
                        
                        if indice >= 30:
                            break
                            
                    if not valores:
                        continue
                
                    for valor in valores:
                        # Guarda as infos da empresa
                        emissao = valor[0]
                        numero = valor[1]
                        cnpj_destinatario = valor[2]
                        nome_destinatario = valor[3]
                        iss = valor[11]
                        pis = valor[5]
                        cofins = valor[8]
                        csll = valor[9]
                        inss = valor[6]
                        irrf = valor[7]
                        nome_destinatario = nome_destinatario.replace('–', '-')
                        _escreve_relatorio_csv(f"{cnpj};{empresa};{emissao};{numero};{nome_destinatario};{cnpj_destinatario};{iss};{pis};{cofins};{csll};{inss};{irrf}", nome='Resumo Reltórios Analíticos')
                except():
                    print(f'\nArquivo: {arq} - ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv(';'.join(['CNPJ', 'NOME', 'EMISSÃO DA NOTA', 'NÚMERO DA NOTA', 'DESTINATÁRIO', 'CNPJ DESTINATÁRIO', 'ISS', 'PIS', 'COFINS', 'CSLL', 'INSS', 'IRRF']), nome='Resumo Reltórios Analíticos')
    print(datetime.now() - inicio)
