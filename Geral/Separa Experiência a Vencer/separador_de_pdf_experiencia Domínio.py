# -*- coding: utf-8 -*-
import fitz, re, os, time
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from pathlib import Path
from datetime import datetime


def ask_for_file(title='Selecione o Relatório de Experiência do Domínio', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    filetypes = [('Plain text files', '*.pdf')]
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    return file if file else False


def ask_for_dir(title='Selecione o local para criar a pasta com os arquivos separados'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False


def escreve_relatorio(pasta_final, andamento):
    try:
        f = open(os.path.join(pasta_final, 'Erros.csv'), 'a', encoding='latin-1')
    except:
        f = open(os.path.join(pasta_final, 'Erros - backup.csv'), 'a', encoding='latin-1')
    
    f.write(f'{andamento};Erro ao coletar informações da empresa no arquivo, layout da página diferente\n')
    f.close()


# guarda info da página anterior
def guarda_info(page, matchtexto_nome, matchtexto_cod):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_cod = matchtexto_cod
    return prevpagina, prevtexto_nome, prevtexto_cod


def cria_pdf(pasta_final, page, matchtexto_nome, matchtexto_cod, prevtexto_nome, pdf, pagina1, pagina2):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome = prevtexto_nome.replace('/', ' ').replace(',', ' ')
        
        text = f'Relatório de Experiências Domínio Web - {data} - {nome}.pdf'
        
        # Define o caminho para salvar o pdf
        os.makedirs(os.path.join(pasta_final, f'Relatorios {data}'), exist_ok=True)
        arquivo = os.path.join(pasta_final, f'Relatorios {data}', text)
        
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)
        
        new_pdf.save(arquivo)
        
        # atualiza as infos da página anterior
        prevpagina = page.number
        prevtexto_nome = matchtexto_nome
        prevtexto_cod = matchtexto_cod
        return prevpagina, prevtexto_nome, prevtexto_cod


def separa():
    # Abrir o pdf
    file = ask_for_file()
    if not file:
        return False
    folder = ask_for_dir()
    if not folder:
        return False
    pasta_final = os.path.join(folder, 'Relatórios de experiência separados Domínio Web')
    os.makedirs(pasta_final, exist_ok=True)
    
    with fitz.open(file) as pdf:
        # Definir os padrões de regex
        padraozinho_nome = re.compile(r'Local\n(.+) - (.+)\n')
        prevpagina = 0
        paginas = 0
        
        # para cada página do pdf
        for page in pdf:
            andamento = f'Pagina = {str(page.number + 1)}'
            '''try:'''
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
            matchtexto_nome = matchzinho_nome.group(2)
            # Guardar o código da empresa no DPCUCA
            matchtexto_cod = matchzinho_nome.group(1)
            
            # Se estiver na primeira página, guarda as informações
            if page.number == 0:
                prevpagina, prevtexto_nome, prevtexto_cod = guarda_info(page, matchtexto_nome, matchtexto_cod)
                continue
            
            # Se o nome da página atual for igual ao da anterior, soma um indice de páginas
            if matchtexto_nome == prevtexto_nome:
                paginas += 1
                # Guarda as informações da página atual
                prevpagina, prevtexto_nome, prevtexto_cod = guarda_info(page, matchtexto_nome, matchtexto_cod)
                continue
            
            # Se for diferente ele separa a página
            else:
                if paginas > 0:
                    # define qual é a primeira página e o nome da empresa
                    paginainicial = prevpagina - paginas
                    andamento = f'Paginas = {str(paginainicial + 1)} até {str(prevpagina + 1)}'
                    prevpagina, prevtexto_nome, prevtexto_cod = cria_pdf(pasta_final, page, matchtexto_nome, matchtexto_cod,
                                                                         prevtexto_nome, pdf,
                                                                         paginainicial, prevpagina)
                    paginas = 0
                # Se for uma página entra a qui
                elif paginas == 0:
                    andamento = f'Pagina = {str(prevpagina + 1)}'
                    prevpagina, prevtexto_nome, prevtexto_cod = cria_pdf(pasta_final, page, matchtexto_nome, matchtexto_cod,
                                                                         prevtexto_nome, pdf,
                                                                         prevpagina, prevpagina)
            '''except:
                escreve_relatorio(pasta_final, andamento)
                continue'''
        
        # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
        try:
            if paginas > 0:
                paginainicial = prevpagina - paginas
                andamento = f'Paginas = {str(paginainicial + 1)} até {str(prevpagina + 1)}'
                cria_pdf(pasta_final, page, matchtexto_nome, matchtexto_cod, prevtexto_nome, pdf, paginainicial,
                         prevpagina)
            elif paginas == 0:
                andamento = f'Pagina = {str(prevpagina + 1)}'
                cria_pdf(pasta_final, page, matchtexto_nome, matchtexto_cod, prevtexto_nome, pdf, prevpagina,
                         prevpagina)
        except:
            escreve_relatorio(pasta_final, andamento)


if __name__ == '__main__':
    mes = datetime.now().strftime('%m')
    ano = datetime.now().strftime('%Y')
    data = f'{mes}-{ano}'
    separa()
    
    messagebox.showinfo(title=None, message='Separador finalizado')
