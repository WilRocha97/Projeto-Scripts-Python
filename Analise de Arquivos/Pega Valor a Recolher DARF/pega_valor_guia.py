# -*- coding: utf-8 -*-
import fitz, re, os
from datetime import datetime

# Definir os padrões de regex
padrao_valor = re.compile(r'Valor Total do Documento\n(.+)')
padrao_info = re.compile(r'Documento de Arrecadação\nde Receitas Federais\n(.+)\n(.+)')
padrao_info2 = re.compile(r'Documento de Arrecadação\ndo eSocial\n(.+)\n(.+)')

# Caminho dos arquivos que serão analizados
arquivos = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', 'Pega Valor a Recolher DARF', 'Guias')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Guias.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Guias - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')


def analiza():
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(arquivos):
        print('\nArquivo: {}'.format(arq))
        # Abrir o pdf
        with fitz.open(r'Guias\{}'.format(arq)) as pdf:

            # Pra cada pagina do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.getText('text', flags=1 + 2 + 8)
                    # Procura o valor a recolher da empresa
                    localiza_valor = padrao_valor.search(textinho)
                    if not localiza_valor:
                        continue

                    # Procura as infos da empresa
                    localiza_info = padrao_info.search(textinho)
                    if not localiza_info:
                        localiza_info = padrao_info2.search(textinho)

                    # Guarda as infos da empresa
                    cnpj = localiza_info.group(1)
                    nome = localiza_info.group(2)
                    valor = localiza_valor.group(1)

                    print('{} - {} - R${}\n'.format(cnpj, nome, valor))

                    texto = ';'.join([cnpj, nome, valor])
                    escrever(texto)

                except():
                    print('ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
