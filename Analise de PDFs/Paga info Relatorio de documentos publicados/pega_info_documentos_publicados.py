# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def analiza():
    # pergunta qual PDF analizar
    relatorio = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    
    # Abrir o pdf
    with fitz.open(relatorio) as pdf:
        empresa_anterior = ''
        # Para cada página do pdf
        for page in pdf:
            '''try:'''
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)
            
            """if page.number == 13:
                print(textinho)
                time.sleep(22)"""
            
            if empresa_anterior != '':
                for i in range(100):
                    recibos = re.compile(r'Data\nUsuário\nData\n(.+\n){' + str(i) + '}').search(textinho)
                    if not recibos:
                        continue
                    # print(recibos)
                    else:
                        if recibos.group(1) == 'Recibo pagamento adiantamento\n':
                            print(empresa_anterior)
                            _escreve_relatorio_csv(empresa_anterior)
                            break
                            
                        if recibos.group(1) == 'Cliente:\n':
                            break
            
            empresas = re.compile(r'(.+)\nCliente:').findall(textinho)
            
            for empresa in empresas:
                for i in range(100):
                    recibos = re.compile(r'' + str(empresa) + '\nCliente:\n(.+\n){' + str(i) + '}').search(textinho)
                    if not recibos:
                        pass
                    # print(recibos)
                    else:
                        if recibos.group(1) == 'Recibo pagamento adiantamento\n':
                            print(empresa)
                            _escreve_relatorio_csv(empresa)
                            break
                        elif recibos.group(1) == 'Cliente:\n':
                            break
                
                empresa_anterior = empresa
            
            '''except:
                _escreve_relatorio_csv(f"{page.number};Erro", nome='erros')'''
        

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()