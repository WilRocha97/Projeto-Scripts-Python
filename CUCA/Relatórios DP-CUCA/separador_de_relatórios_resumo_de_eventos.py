# -*- coding: utf-8 -*-
import fitz
import re
import os
import pandas as pd
import time
from pathlib import Path
from datetime import datetime
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory, Tk


def ask_for_dir(title=''):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(title=title)
    
    return folder if folder else False


def guarda_info(page, resumo_evento):
    prevpagina = page.number
    prev_resumo_evento = resumo_evento
    return prevpagina, prev_resumo_evento


def cria_pdf(page, resumo_evento, pdf, pasta_final, file_name, paginainicial, prevpagina, andamento):
    with fitz.open() as new_pdf:
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=paginainicial, to_page=prevpagina)
        print(file_name)
        new_pdf.save(os.path.join(pasta_final, file_name))
        print(andamento)

        prevpagina = page.number
        prev_resumo_evento = resumo_evento
        return prevpagina, prev_resumo_evento


def separa():
    # Abrir o pdf
    documentos = ask_for_dir(title='Selecione a pasta onde estão os PDFs para separar')
    if not documentos:
        return False
    folder = ask_for_dir(title='Selecione o local para criar a pasta com os arquivos separados')
    if not folder:
        return False
    pasta_final = os.path.join(folder, 'Resumos por evento')
    os.makedirs(pasta_final, exist_ok=True)
    
    messagebox.showinfo(title=None, message='Locais selecionados, clique em "OK" e aguarde o procedimento finalizar.')
    
    for file_name in os.listdir(documentos):
        print(file_name)
        file = os.path.join(documentos, file_name)
        try:
            with fitz.open(file) as pdf:
                pdf = 'erro'
        except:
            continue
        # Abrir o pdf
        with fitz.open(file) as pdf:
            prevpagina = 0
            paginas = 0
            resumo_evento = ''
            prev_resumo_evento = ''
            
            # para cada página do pdf
            for page in pdf:
                # try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                try:
                    resumo_evento = re.compile(r'(Resumo por Eventos)').search(textinho).group(1)
                except:
                    # print(textinho)
                    continue

                # Se estiver na primeira página, guarda as informações
                if page.number == 0:
                    prevpagina, prev_resumo_evento = guarda_info(page, resumo_evento)
                    continue

                # Se o nome da página atual for igual ao da anterior, soma um indice de páginas
                if resumo_evento == prev_resumo_evento:
                    paginas += 1
                    # Guarda as informações da página atual
                    prevpagina, prev_resumo_evento = guarda_info(page, resumo_evento)
                    continue

                # Se for diferente ele separa a página
                else:
                    # Se for mais de uma página entra aqui
                    if paginas > 0:
                        # define qual é a primeira página e o nome da empresa
                        paginainicial = prevpagina - paginas
                        andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prev_resumo_evento = cria_pdf(page, resumo_evento, pdf, pasta_final, file_name, paginainicial, prevpagina, andamento)
                        paginas = 0
                    # Se for uma página entra a qui
                    elif paginas == 0:
                        andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                        prevpagina, prev_resumo_evento = cria_pdf(page, resumo_evento, pdf, pasta_final, file_name, prevpagina, prevpagina, andamento)
                '''except:
                    print('❌ ERRO')
                    continue'''
    
            # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
            if paginas > 0:
                paginainicial = prevpagina - paginas
                andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                cria_pdf(page, resumo_evento, pdf, pasta_final, file_name, paginainicial, prevpagina, andamento)
                
            elif paginas == 0:
                andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                cria_pdf(page, resumo_evento, pdf, pasta_final, file_name, prevpagina, prevpagina, andamento)


if __name__ == '__main__':
    # o robo pega o pdf na pasta PDF e cria outro para colocar os separados
    inicio = datetime.now()
    separa()
    print(datetime.now() - inicio)
