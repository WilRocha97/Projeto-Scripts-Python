# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

# Definir os padrões de regex
padrao_empresa = re.compile(r'CNPJ Matriz:\n(.+)\nNome empresarial:\n(.+)')
padrao_declaracao = re.compile(r'Nº da Declaração:\n(.+)')
padrao_receita = re.compile(r'Receita bruta acumulada no ano-calendário corrente\n\(RBA\)\n(.+)\n(.+)\n(.+)')

# Caminho dos arquivos que serão analizados
arquivos = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', 'Pega Valor a Recolher DARF', 'PDF')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Resumo Mensal.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Resumo Mensal - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')


def analiza():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:

            # Pra cada pagina do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # print(textinho)
                    # time.sleep(55)
                    # Procura o valor a recolher da empresa
                    localiza_empresa = padrao_empresa.search(textinho)

                    if not localiza_empresa:
                        continue

                    # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a éssa info
                    localiza_receita = padrao_receita.search(textinho)
                    localiza_declaracao = padrao_declaracao.search(textinho)
                    
                    # Guarda as infos da empresa
                    cnpj = localiza_empresa.group(1)
                    nome = localiza_empresa.group(2)
                    declaracao = localiza_declaracao.group(1)
                    mercado_interno = localiza_receita.group(1)
                    mercado_externo = localiza_receita.group(2)
                    total = localiza_receita.group(3)
                    
                    print(f'{cnpj} - {nome}')

                    _escreve_relatorio_csv(f"{cnpj};{nome};{mercado_interno};{mercado_externo};{total};Nº {declaracao}")

                except():
                    print('ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv(';'.join(['CNPJ', 'NOME', 'MERCADO INTERNO', 'MERCADO EXTERNO', 'TOTAL', 'NÚMERO DA DECLARAÇÃO']))
    print(datetime.now() - inicio)
