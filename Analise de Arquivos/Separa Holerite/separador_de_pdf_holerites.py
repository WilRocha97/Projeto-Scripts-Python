# -*- coding: utf-8 -*-
import fitz, re, os, pyautogui as p
from datetime import datetime
from tkinter import messagebox
from pathlib import Path
from tkinter.filedialog import askopenfilename, askdirectory, Tk


def ask_for_dir(title=''):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(title=title)
    
    return folder if folder else False


def guarda_info(page, matchtexto_nome):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, pagina1, pagina2, andamento, file_name):
    if prevtexto_nome == quem:
        with fitz.open() as new_pdf:
            # Define o nome do arquivo
            file_name = file_name.replace('.PDF', '').replace('.pdf', '')
            nome = prevtexto_nome
            nome = nome.replace('/', ' ')
            text = '{} - {}.pdf'.format(file_name, nome)

            # Define o caminho para salvar o pdf
            text = os.path.join('Holerites', text)

            # Define a página inicial e a final
            new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)

            new_pdf.save(text)
            print(nome + andamento)

    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def separa():
    # Abrir o pdf
    documentos = ask_for_dir(title='Selecione a pasta onde estão os PDFs para separar')
    if not documentos:
        return False
    
    # Para cada arquivo nos dados
    for file_name in os.listdir(documentos):
        file = os.path.join(documentos, file_name)
        print(file_name)
        # Abrir o pdf
        with fitz.open(file) as pdf:
            # Definir os padrões de regex
            padraozinho_nome = re.compile(r'(.+)\nBanco:')
            padraozinho_nome_2 = re.compile(r'(.+)\nPIS:')
            prevpagina = 0
            paginas = 0
            
            # Para cada página do pdf
            for page in pdf:
                
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    
                    # Procura o nome do funcionario
                    matchzinho_nome = padraozinho_nome_2.search(textinho)

                    if not matchzinho_nome:
                        prevpagina = page.number
                        continue

                    # Guardar o nome da empresa
                    matchtexto_nome = matchzinho_nome.group(1)

                    # Se estiver na primeira página, guarda as informações
                    if page.number == 0:
                        prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                        continue

                    # Se o nome da página atual for igual ao da anterior, soma mais um no indice de páginas
                    if matchtexto_nome == prevtexto_nome:
                        paginas += 1
                        # Guarda as informações da página atual
                        prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                        continue

                    # Se for diferente ele separa a página
                    else:
                        # Se for mais de uma página entra aqui
                        if paginas > 0:
                            # define qual é a primeira página e o nome da empresa
                            paginainicial = prevpagina - paginas
                            andamento = ' - Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1)
                            prevpagina, prevtexto_nome = cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, paginainicial, prevpagina, andamento, file_name)
                            paginas = 0
                        # Se for uma página entra a qui
                        elif paginas == 0:
                            andamento = ' - Pagina = ' + str(prevpagina + 1)
                            prevpagina, prevtexto_nome = cria_pdf(page, matchtexto_nome,  prevtexto_nome, pdf, prevpagina, prevpagina, andamento, file_name)
                except:
                    print('❌ ERRO')
                    continue

            # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
            if paginas > 0:
                paginainicial = prevpagina - paginas
                andamento = ' - Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1)
                cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, paginainicial, prevpagina, andamento, file_name)
            elif paginas == 0:
                andamento = ' - Pagina = ' + str(prevpagina + 1)
                cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, prevpagina, prevpagina, andamento, file_name)


if __name__ == '__main__':
    os.makedirs('Holerites', exist_ok=True)
    inicio = datetime.now()

    quem = p.prompt(text='Qual funcionário?\n(Escreva exatamente como está no Holerite)', title='Script incrível')
    separa()

    print(datetime.now() - inicio)
