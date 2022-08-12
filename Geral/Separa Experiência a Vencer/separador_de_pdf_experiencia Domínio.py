# -*- coding: utf-8 -*-
import fitz, re, os, time
from pathlib import Path
from datetime import datetime


def guarda_info(page, matchtexto_nome, matchtexto_cnpj):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_cnpj = matchtexto_cnpj
    return prevpagina, prevtexto_nome, prevtexto_cnpj


def cria_pdf(page, matchtexto_nome, matchtexto_cnpj, prevtexto_nome, prevtexto_cnpj, pdf, pagina1, pagina2, andamento):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome = prevtexto_nome
        cnpj = prevtexto_cnpj.replace('.', '').replace('/', '').replace('-', '')
        text = cnpj + ';EXPERIENCIA;' + mes + '-' + ano + ' ' + nome + '.pdf'
        
        # Define o caminho para salvar o pdf
        text = os.path.join('Separados Domínio', text)
        
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)
        
        new_pdf.save(text)
        print(nome + andamento)
        
        prevpagina = page.number
        prevtexto_nome = matchtexto_nome
        prevtexto_cnpj = matchtexto_cnpj
        return prevpagina, prevtexto_nome, prevtexto_cnpj


def separa():
    # Abrir o pdf
    with fitz.open(r'PDF\Experiência Domínio.pdf') as pdf:
        # Definir os padrões de regex
        padraozinho_nome = re.compile(r'Empresa: \d+ - (.+)\n.+: (\d.+)\nPágina')
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
                matchtexto_nome = matchzinho_nome.group(1)
                # Guardar o código da empresa no DPCUCA
                matchtexto_cnpj = matchzinho_nome.group(2)
                
                # Se estiver na primeira página, guarda as informações
                if page.number == 0:
                    prevpagina, prevtexto_nome, prevtexto_cnpj = guarda_info(page, matchtexto_nome, matchtexto_cnpj)
                    continue
                
                # Se o nome da página atual for igual ao da anterior, soma um indice de páginas
                if matchtexto_nome == prevtexto_nome:
                    paginas += 1
                    # Guarda as informações da página atual
                    prevpagina, prevtexto_nome, prevtexto_cnpj = guarda_info(page, matchtexto_nome, matchtexto_cnpj)
                    continue
                
                # Se for diferente ele separa a página
                else:
                    # Se for mais de uma página entra aqui
                    if paginas > 0:
                        # define qual é a primeira página e o nome da empresa
                        paginainicial = prevpagina - paginas
                        andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_cnpj = cria_pdf(page, matchtexto_nome, matchtexto_cnpj,
                                                                              prevtexto_nome, prevtexto_cnpj, pdf,
                                                                              paginainicial, prevpagina, andamento)
                        paginas = 0
                    # Se for uma página entra a qui
                    elif paginas == 0:
                        andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prevtexto_nome, prevtexto_cnpj = cria_pdf(page, matchtexto_nome, matchtexto_cnpj,
                                                                              prevtexto_nome, prevtexto_cnpj, pdf,
                                                                              prevpagina, prevpagina, andamento)
            except:
                print('❌ ERRO')
                continue
        
        # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
        if paginas > 0:
            paginainicial = prevpagina - paginas
            andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_cnpj, prevtexto_nome, prevtexto_cnpj, pdf, paginainicial,
                     prevpagina, andamento)
        elif paginas == 0:
            andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
            cria_pdf(page, matchtexto_nome, matchtexto_cnpj, prevtexto_nome, prevtexto_cnpj, pdf, prevpagina,
                     prevpagina, andamento)


if __name__ == '__main__':
    # o robo pega o pdf na pasta PDF e cria outro para colocar os separados
    os.makedirs('Separados', exist_ok=True)
    inicio = datetime.now()

    # Definir o mês e o ano da consulta
    mes = datetime.now().strftime('%m')
    ano = datetime.now().strftime('%Y')
    
    separa()
    
    print(datetime.now() - inicio)
