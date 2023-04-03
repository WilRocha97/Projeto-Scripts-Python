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

                total_servico = re.compile(r'Totais\n.+,.+\n(.+,.+)').search(textinho).group(1)
                nome = re.compile(r'Empresa:\n(.+)').search(textinho).group(1)
                cnpj = re.compile(r'CNPJ:\n(.+)').search(textinho).group(1)
                
                periodo = re.compile(r'Período:\n(.+\n.+)').search(textinho).group(1)
                periodo = periodo.replace('/', '-').replace('\na', ' até ')
                
                _escreve_relatorio_csv(f"{cnpj};{nome};{periodo};{total_servico}", nome=f'Relatório de Faturamento Serviço - {periodo}')
        
            except():
                print(textinho)
                print('ERRO')

    return periodo


if __name__ == '__main__':
    inicio = datetime.now()
    periodo = analiza()
    _escreve_header_csv('CNPJ;NOME;PERIODO;TOTAL DE SERVIÇOS', nome=f'Relatório de Faturamento Serviço - {periodo}')
    print(datetime.now() - inicio)
