# -*- coding: utf-8 -*-
import fitz
import re
import sys
import os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file

def guarda_info(page, matchtexto_nome):
    prevpagina = page.number
    # prevtexto_cnpj = matchtexto
    # prevtexto_codigo = matchtexto2
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def cria_pdf(pdf, page, matchtexto_nome, prevtexto_nome, pagina1, pagina2, andamento):
    # prevtexto_cnpj = prevtexto.replace('/', '').replace('.', '').replace('-', '')
    prevtexto_nome = prevtexto_nome.replace('/', '')
    with fitz.open() as new_pdf:
        text = prevtexto_nome + ' - Informe de Rendimento Funcionários.pdf'
        
        # Define o caminho para salvar o pdf
        text = os.path.join('Separados', text)
        
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)
        new_pdf.save(text)
    print(andamento)
    prevpagina = page.number
    # prevtexto_cnpj = matchtexto
    # prevtexto_codigo = matchtexto2
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def separa():
    # Abrir o pdf
    file = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    if not file:
        return False
    # Abrir o PDF
    with fitz.open(file) as pdf:
        # padraozinho_cnpj = re.compile(r'(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})')
        # padraozinho_codigo = re.compile(r'(\d{5})')
        padraozinho_nome = re.compile(r'Responsável pelas Informações\n(.+)')
        padraozinho_nome2 = re.compile(r'Fax:\n(.+)')
        prevpagina = 0
        paginas = 0

        for page in pdf:
            textinho = page.get_text('text', flags=1+2+8)
            # matchzinho_cnpj = padraozinho.search(textinho)
            # matchzinho_codigo = padraozinho2.search(textinho)
            matchzinho_nome = padraozinho_nome.search(textinho)
            if not matchzinho_nome:
                matchzinho_nome = padraozinho_nome2.search(textinho)
                if not matchzinho_nome:
                    prevpagina = page.number
                    continue

            # matchtexto_cnpj = matchzinho.group(1)
            # matchtexto_codigo = matchzinho2.group(1)
            matchtexto_nome = matchzinho_nome.group(1)

            if page.number == 0:
                prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                continue

            if matchtexto_nome == prevtexto_nome:
                paginas += 1
                prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                continue

            else:
                if paginas > 0:
                    paginainicial = prevpagina - paginas
                    andamento = prevtexto_nome + '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                    prevpagina, prevtexto_nome = cria_pdf(pdf, page, matchtexto_nome, prevtexto_nome, paginainicial, prevpagina, andamento)
                    paginas = 0

                elif paginas == 0:
                    andamento = prevtexto_nome + '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                    prevpagina, prevtexto_nome = cria_pdf(pdf, page, matchtexto_nome, prevtexto_nome, prevpagina, prevpagina, andamento)

        if paginas > 0:
            paginainicial = prevpagina - paginas
            andamento = prevtexto_nome + '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(pdf, page, matchtexto_nome, prevtexto_nome, paginainicial, prevpagina, andamento)

        elif paginas == 0:
            andamento = prevtexto_nome + '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(pdf, page, matchtexto_nome, prevtexto_nome, paginainicial, prevpagina, andamento)

    # cnpj_cru = "00.000.000/0000-00"
    # cnpj = [i for i in cnpj_cru if i.isdigit()]


if __name__ == '__main__':
    inicio = datetime.now()

    separa()

    print(datetime.now() - inicio)
