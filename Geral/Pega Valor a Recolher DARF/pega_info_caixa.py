# -*- coding: utf-8 -*-
import fitz, re, os, time
from datetime import datetime

# Definir os padrões de regex
padrao_cnpj = re.compile(r'Inscrição:\n(.+)')
padrao_nome = re.compile(r'Razão	Social:\n(.+)')
padrao_endereco = re.compile(r'Endereço:\n(.+)')

# Caminho dos arquivos que serão analizados
arquivos = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', 'Pega Valor a Recolher DARF', 'Guias')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Caixa.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Caixa - Auxiliar.csv', 'a') as f:
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
                    localiza_cnpj = padrao_cnpj.search(textinho)
                    if not localiza_cnpj:
                        continue

                    # Procura as infos da empresa
                    localiza_nome = padrao_nome.search(textinho)
                    localiza_endereco = padrao_endereco.search(textinho)

                    # Guarda as infos da empresa
                    cnpj = localiza_cnpj.group(1)
                    nome = localiza_nome.group(1)
                    endereco = localiza_endereco.group(1)

                    print('{} - {} - R${}\n'.format(cnpj, nome, endereco))

                    texto = ';'.join([cnpj, nome, endereco])
                    escrever(texto)

                except():
                    print('ERRO')


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
