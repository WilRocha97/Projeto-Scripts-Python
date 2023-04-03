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

                total_entrada = re.compile(r'Total de Entradas:\n(.+,\d\d)').search(textinho).group(1)
                total_saida = re.compile(r'Total de Saídas:\n(.+,\d\d)').search(textinho).group(1)
                resultado = re.compile(r'Resultado:\n(.+,\d\d)').search(textinho).group(1)
                porcentagem = re.compile(r'Porcentagem:\n(.+%)').search(textinho).group(1)
                empresa = re.compile(r'(.+)\n.+\n.+\n.+\n.+\n.+\n.+\nCNPJ:\n.+\n(.+)\nPeríodo:\n(.+)').search(textinho)
                
                cnpj = empresa.group(2)
                nome = empresa.group(1)
                periodo = empresa.group(3)
                periodo = periodo.replace('/', '-')
                
                _escreve_relatorio_csv(f"{cnpj};{nome};{periodo};{total_entrada};{total_saida};{resultado};{porcentagem}", nome=f'Comparativo de Entradas e Saídas - {periodo}')
        
            except():
                print(textinho)
                print('ERRO')

    return periodo


if __name__ == '__main__':
    inicio = datetime.now()
    periodo = analiza()
    _escreve_header_csv('CNPJ;NOME;PERIODO;TOTAL DE ENTRADA;TOTAL DE SAÍDA;RESULTADO;PORCENTAGEM', nome=f'Comparativo de Entradas e Saídas - {periodo}')
    print(datetime.now() - inicio)
