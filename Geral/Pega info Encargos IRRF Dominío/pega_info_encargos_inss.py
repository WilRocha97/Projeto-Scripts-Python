# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

# Definir os padrões de regex
padrao_cnpj = re.compile(r'(\d+) - (.+)\n(.+)')


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

                # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a essa info

                valores = re.compile(r'(\d\d\d\d)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(\d\d/\d\d\d\d)\n(.+)').findall(textinho)
                if not valores:
                    continue
                    
                for valor in valores:
                    empresa = re.compile(r'(\d+) - (.+)\n(.+)').search(textinho)
                    
                    codigo = empresa.group(1)
                    nome = empresa.group(2)
                    cnpj = empresa.group(3)
                    
                    periodicidade = valor[6]
                    periodo = valor[5]
                    codigo_receita = valor[0]
                    valor_recolher = valor[1]
                    valor_compensar = valor[2]
                    valor_pagar = valor[3]
                    valor_acumular = valor[4]
                    
                    # print(f'{codigo} - {cnpj} - {nome} - {valor}')
    
                    _escreve_relatorio_csv(f"{codigo};{cnpj};{nome};{periodicidade};{periodo};{codigo_receita};{valor_recolher};{valor_compensar};{valor_pagar};{valor_acumular}", nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
        
            except():
                print(textinho)
                print('ERRO')
                
    return periodo
    

if __name__ == '__main__':
    inicio = datetime.now()
    periodo = analiza()
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;PERIODICIDADE;PERÍODO;CÓDIGO RECOLHIMENTO;VALOR A RECOLHER;VALOR A COMPENSAR;VALOR A PAGAR;VALOR A ACUMULAR', nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
    print(datetime.now() - inicio)
