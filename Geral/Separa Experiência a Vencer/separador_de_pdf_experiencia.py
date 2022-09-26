# -*- coding: utf-8 -*-
import fitz
import re
import os
import pandas as pd
import time
from pathlib import Path
from datetime import datetime


def configura():
    # Definir o mês e o ano da consulta
    mes_consulta = datetime.now().strftime('%m')
    ano_consulta = datetime.now().strftime('%Y')

    # Pega a planilha mais recente
    data_criacao = (lambda f: f.stat().st_ctime)
    directory = Path('V:/Setor Robô/Relatorios/dp')
    files = directory.glob('*.csv')
    sorted_files = sorted(files, key=data_criacao, reverse=True)
    dados = str(sorted_files[0]).split(sep='\\')
    dados = dados[4]

    # Definir o nome das colunas
    colunas = ['cod', 'cnpj', 'nome']
    # Localiza a planilha
    caminho = Path('V:/Setor Robô/Relatorios/dp/' + dados)
    # Abrir a planilha
    lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
    # Definir o index da planilha
    lista.set_index('cod', inplace=True)

    return lista, mes_consulta, ano_consulta


def guarda_info(page, matchtexto_nome, matchtexto_codigo):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_codigo = matchtexto_codigo
    return prevpagina, prevtexto_nome, prevtexto_codigo


def pegar_cnpj(prevtexto_codigo):
    # Procura com o código da empresa para pegar o cnpj
    cnpj = arquivo.loc[int(prevtexto_codigo)]
    cnpj = list(cnpj)
    cnpj = cnpj[0]
    cnpj = cnpj.replace('/', '').replace('-', '').replace('.', '').replace(' ', '')
    return cnpj


def cria_pdf(page, matchtexto_nome, matchtexto_codigo, prevtexto_nome, prevtexto_codigo, pdf, pagina1, pagina2, andamento):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome = prevtexto_nome
        cnpj = pegar_cnpj(prevtexto_codigo)
        text = cnpj + ';EXPERIENCIA;' + mes + '-' + ano + ' ' + nome + '.pdf'

        # Define o caminho para salvar o pdf
        text = os.path.join('Separados', text)

        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)

        new_pdf.save(text)
        print(nome + andamento)

        prevpagina = page.number
        prevtexto_nome = matchtexto_nome
        prevtexto_codigo = matchtexto_codigo
        return prevpagina, prevtexto_nome, prevtexto_codigo


def separa():
    # Abrir o pdf
    with fitz.open(r'PDF\Experiência.pdf') as pdf:
        # Definir os padrões de regex
        padraozinho_nome = re.compile(r'a Vencer+(\n.+){7}\n0+(\d{1,4})(.+)')
        padraozinho_nome2 = re.compile(r'a Vencer(\n.+){7}\n(\d{1,4})(.+)')
        prevpagina = 0
        paginas = 0

        # para cada página do pdf
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                matchzinho_nome = padraozinho_nome.search(textinho)
                if not matchzinho_nome:
                    matchzinho_nome = padraozinho_nome2.search(textinho)
                    if not matchzinho_nome:
                        prevpagina = page.number
                        continue

                # Guardar o nome da empresa
                matchtexto_nome = matchzinho_nome.group(3)
                # Guardar o código da empresa no DPCUCA
                matchtexto_codigo = matchzinho_nome.group(2)

                # Se estiver na primeira página, guarda as informações
                if page.number == 0:
                    prevpagina, prevtexto_nome, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_codigo)
                    continue

                # Se o nome da página atual for igual ao da anterior, soma um indice de páginas
                if matchtexto_nome == prevtexto_nome:
                    paginas += 1
                    # Guarda as informações da página atual
                    prevpagina, prevtexto_nome, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_codigo)
                    continue

                # Se for diferente ele separa a página
                else:
                    # Se for mais de uma página entra aqui
                    if paginas > 0:
                        # define qual é a primeira página e o nome da empresa
                        paginainicial = prevpagina - paginas
                        andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_codigo = cria_pdf(page, matchtexto_nome, matchtexto_codigo,
                                                                                prevtexto_nome, prevtexto_codigo, pdf,
                                                                                paginainicial, prevpagina, andamento)
                        paginas = 0
                    # Se for uma página entra a qui
                    elif paginas == 0:
                        andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_codigo = cria_pdf(page, matchtexto_nome, matchtexto_codigo,
                                                                                prevtexto_nome, prevtexto_codigo, pdf,
                                                                                prevpagina, prevpagina, andamento)
            except:
                print('❌ ERRO')
                continue

        # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
        if paginas > 0:
            paginainicial = prevpagina - paginas
            andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_codigo, prevtexto_nome, prevtexto_codigo, pdf, paginainicial,
                     prevpagina, andamento)
            
        elif paginas == 0:
            andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_codigo, prevtexto_nome, prevtexto_codigo, pdf, prevpagina,
                     prevpagina, andamento)


if __name__ == '__main__':
    # o robo pega o pdf na pasta PDF e cria outro para colocar os separados
    os.makedirs('Separados', exist_ok=True)
    inicio = datetime.now()

    arquivo, mes, ano = configura()
    separa()

    print(datetime.now() - inicio)
