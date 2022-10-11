# -*- coding: utf-8 -*-
import fitz
import re
import os
from datetime import datetime


def guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_cnpj = matchtexto_cnpj
    prevtexto_codigo = matchtexto_codigo
    return prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo


def cria_pdf(page, matchtexto_nome, matchtexto_cnpj, matchteexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, pagina1, pagina2, andamento):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome = prevtexto_nome
        nome = nome.replace('/', ' ').replace(',', '')
        cnpj = prevtexto_cnpj
        cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
        codigo = prevtexto_codigo
        text = '{} - {} - {}.pdf'.format(nome, cnpj, codigo)

        # Define o caminho para salvar o pdf
        text = os.path.join('Separados Relatórios', text)

        # Define a página inicial e a final
        new_pdf.insertPDF(pdf, from_page=pagina1, to_page=pagina2)

        new_pdf.save(text)
        print(nome + andamento)

        prevpagina = page.number
        prevtexto_nome = matchtexto_nome
        prevtexto_cnpj = matchtexto_cnpj
        prevtexto_codigo = matchteexto_codigo
        return prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo


def separa():
    # Abrir o pdf
    with fitz.open(r'PDFs\Relatórios.pdf') as pdf:
        # Definir os padrões de regex
        padraozinho_cnpj = re.compile(r'(.+)\nCNPJ')
        padraozinho_nome = re.compile(r'F.Social:\n(.+)')
        padraozinho_codigo = re.compile(r'Código:(.+)')
        prevpagina = 0
        paginas = 0

        # Pra cada pagina do pdf
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.getText('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                matchzinho_nome = padraozinho_nome.search(textinho)
                matchzinho_cnpj = padraozinho_cnpj.search(textinho)
                matchzinho_codigo = padraozinho_codigo.search(textinho)
                if not matchzinho_nome:
                    matchzinho_nome = padraozinho_nome.search(textinho)
                    matchzinho_cnpj = padraozinho_cnpj.search(textinho)
                    matchzinho_codigo = padraozinho_codigo.search(textinho)
                    if not matchzinho_nome:
                        prevpagina = page.number
                        continue

                # Guardar o nome da empresa
                matchtexto_nome = matchzinho_nome.group(1)
                # Guardar o cnpj da empresa no DPCUCA
                matchtexto_cnpj = matchzinho_cnpj.group(1)
                # Guardar o código da empresa no DPCUCA
                matchtexto_codigo = matchzinho_codigo.group(1)

                # Se estiver na primeira página, guarda as informações
                if page.number == 0:
                    prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo)
                    continue

                # Se o nome da página atual for igual ao da anterior, soma mais um no indice de paginas
                if matchtexto_nome == prevtexto_nome:
                    paginas += 1
                    # Guarda as informações da pagina atual
                    prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo)
                    continue

                # Se for diferente ele separa a página
                else:
                    # Se for mais de uma página entra aqui
                    if paginas > 0:
                        # define qual é a primeira página e o nome da empresa
                        paginainicial = prevpagina - paginas
                        andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = cria_pdf(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome,
                                                                                                prevtexto_cnpj, prevtexto_codigo, pdf, paginainicial, prevpagina, andamento)
                        paginas = 0
                    # Se for uma página entra a qui
                    elif paginas == 0:
                        andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = cria_pdf(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome,
                                                                                                prevtexto_cnpj, prevtexto_codigo, pdf, prevpagina, prevpagina, andamento)
            except:
                print('❌ ERRO')
                continue

        # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
        if paginas > 0:
            paginainicial = prevpagina - paginas
            andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, paginainicial, prevpagina, andamento)
        elif paginas == 0:
            andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, prevpagina, prevpagina, andamento)


if __name__ == '__main__':
    # o robo pega os pdfs na pasta PDFs e cria uma nova para colocar os separados
    os.makedirs('Separados Relatórios', exist_ok=True)
    inicio = datetime.now()
    separa()

    print(datetime.now() - inicio)
