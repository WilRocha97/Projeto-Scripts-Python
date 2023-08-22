# -*- coding: utf-8 -*-
import fitz, re, os, time, shutil
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from pathlib import Path
from datetime import datetime


def ask_for_dir(title=''):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(title=title)
    
    return folder if folder else False


def escreve_relatorio(pasta_final, andamento):
    try:
        f = open(os.path.join(pasta_final, 'Erros.csv'), 'a', encoding='latin-1')
    except:
        f = open(os.path.join(pasta_final, 'Erros - backup.csv'), 'a', encoding='latin-1')
    
    f.write(f'{andamento};Erro ao coletar informações da empresa no arquivo, layout da página diferente\n')
    f.close()
    
    
def guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    prevtexto_cnpj = matchtexto_cnpj
    prevtexto_codigo = matchtexto_codigo
    return prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo


def cria_pdf(sistema, pasta_final, page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, pagina1, pagina2):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome = prevtexto_nome
        nome = nome.replace('/', ' ')
        cnpj = prevtexto_cnpj
        cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
        codigo = prevtexto_codigo
        text = '{} - {} - {} - DARF {}.pdf'.format(nome, cnpj, codigo, sistema)

        # Define o caminho para salvar o pdf
        os.makedirs(os.path.join(pasta_final, f'DARFs {codigo}'), exist_ok=True)
        arquivo = os.path.join(pasta_final, f'DARFs {codigo}', text)

        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)

        new_pdf.save(arquivo)

        prevpagina = page.number
        prevtexto_nome = matchtexto_nome
        prevtexto_cnpj = matchtexto_cnpj
        prevtexto_codigo = matchtexto_codigo
        return prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo


def separa_duplicatas(pasta_final):
    pasta_repetidas = os.path.join(pasta_final, 'Empresas repetidas')
    for folder in os.listdir(pasta_final):
        for file in os.listdir(os.path.join(pasta_final, folder)):
            repetidas = procura(file, pasta_final, pasta_repetidas)
            if repetidas == 'ok':
                try:
                    shutil.move(os.path.join(pasta_final, folder, file), pasta_repetidas)
                except:
                    pass
            
            
def procura(file, pasta_final, pasta_repetidas):
    repetidas = 'não'
    for folder in os.listdir(pasta_final):
        for file_compara in os.listdir(os.path.join(pasta_final, folder)):
            if file != file_compara:
                file_s = str(file).split('-')
                file_compara_s = str(file_compara).split('-')
                # print(f'{file_s[1]} - {file_compara_s[1]}')
                if file_s[1] == file_compara_s[1] and file_s[2] == file_compara_s[2]:
                    os.makedirs(pasta_repetidas, exist_ok=True)
                    try:
                        shutil.move(os.path.join(pasta_final, folder, file_compara), pasta_repetidas)
                    except:
                        pass
                    repetidas = 'ok'
                    
    return repetidas

                
def separa():
    # Abrir o pdf
    documentos = ask_for_dir(title='Selecione a pasta onde estão os PDFs para separar')
    if not documentos:
        return False
    folder = ask_for_dir(title='Selecione o local para criar a pasta com os arquivos separados')
    if not folder:
        return False
    pasta_final = os.path.join(folder, 'DARFs')
    os.makedirs(pasta_final, exist_ok=True)

    messagebox.showinfo(title=None, message='Locais selecionados, clique em "OK" e aguarde o procedimento finalizar.')
    
    for file in os.listdir(documentos):
        file = os.path.join(documentos, file)
        try:
            with fitz.open(file) as pdf:
                pdf = 'erro'
        except:
            continue
        # Abrir o pdf
        with fitz.open(file) as pdf:
            
            # Definir os padrões de regex
            padraozinho_cnpj = re.compile(r'03 Número de CPF ou CNPJ\n(.+)')
            padraozinho_nome = re.compile(r'06 Data de Vencimento\n.+\n(.+)')
            padraozinho_codigo = re.compile(r'Receitas Federais\n(.+)')
            padraozinho_dominio = re.compile(r'(.+)\n(\d\d\.\d\d\d\.\d\d\d\/\d\d\d\d-\d\d)(\n.+){4}\n(\d\d\d\d)')
            padraozinho_dominio_2 = re.compile(r'(.+)\n(\d\d\.\d\d\d\.\d\d\d\/\d\d\d\d-\d\d)')
            padraozinho_dominio_pis = re.compile(r'DARF\n(PIS)')
            padraozinho_dominio_cpf = re.compile(r'(.+)\n(\d\d\d.\d\d\d.\d\d\d-\d\d)(\n.+){4}\n(\d\d\d\d)')
            prevpagina = 0
            paginas = 0
    
            # para cada página do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # print(textinho)
                    try:
                        try:
                            # Procura o nome da empresa no texto do pdf
                            matchzinho_nome = padraozinho_nome.search(textinho)
                            matchzinho_cnpj = padraozinho_cnpj.search(textinho)
                            matchzinho_codigo = padraozinho_codigo.search(textinho)
                            # Guardar o nome da empresa
                            matchtexto_nome = matchzinho_nome.group(1).replace('-', ' ')
                            # Guardar o cnpj da empresa no DPCUCA
                            matchtexto_cnpj = matchzinho_cnpj.group(1)
                            # Guardar o codigo da empresa no DPCUCA
                            matchtexto_codigo = matchzinho_codigo.group(1).replace(' ', '')
                            sistema = 'CUCA'
                        except:
                            try:
                                matchzinho_nome = padraozinho_dominio.search(textinho)
                                # Guardar o nome da empresa
                                matchtexto_nome = matchzinho_nome.group(1).replace('-', ' ')
                                # Guardar o cnpj da empresa no DPCUCA
                                matchtexto_cnpj = matchzinho_nome.group(2)
                                # Guardar o codigo da empresa no DPCUCA
                                matchtexto_codigo = matchzinho_nome.group(4).replace(' ', '')
                            except:
                                try:
                                    matchzinho_nome = padraozinho_dominio_2.search(textinho)
                                    matchzinho_codigo = padraozinho_dominio_pis.search(textinho)
                                    # Guardar o nome da empresa
                                    matchtexto_nome = matchzinho_nome.group(1).replace('-', ' ')
                                    # Guardar o cnpj da empresa no DPCUCA
                                    matchtexto_cnpj = matchzinho_nome.group(2)
                                    # Guardar o codigo da empresa no DPCUCA
                                    matchtexto_codigo = matchzinho_codigo.group(1).replace(' ', '')
                                except:
                                    matchzinho_nome = padraozinho_dominio_cpf.search(textinho)
                                    # Guardar o nome da empresa
                                    matchtexto_nome = matchzinho_nome.group(1).replace('-', ' ')
                                    # Guardar o cnpj da empresa no DPCUCA
                                    matchtexto_cnpj = matchzinho_nome.group(2)
                                    # Guardar o codigo da empresa no DPCUCA
                                    matchtexto_codigo = matchzinho_nome.group(4).replace(' ', '')
                            sistema = 'DOMÍNIO WEB'
                    except:
                        print(textinho)
                        
                    if not matchzinho_nome:
                        prevpagina = page.number
                        continue
                    
                    # Se estiver na primeira página, guarda as informações
                    if page.number == 0:
                        prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo)
                        continue
        
                    # Se o nome da página atual for igual ao da anterior, soma 1 no indice de páginas
                    if matchtexto_nome == prevtexto_nome and matchtexto_codigo == prevtexto_codigo:
                        paginas += 1
                        # Guarda as informações da página atual
                        prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = guarda_info(page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo)
                        continue
                        
                    # Se for diferente ele separa a página
                    else:
                        # Se for mais de uma página entra aqui
                        if paginas > 0:
                            # define qual é a primeira página e o nome da empresa
                            paginainicial = prevpagina - paginas
                            andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                            prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = cria_pdf(sistema, pasta_final, page, matchtexto_nome, matchtexto_cnpj,
                                                                                                    matchtexto_codigo, prevtexto_nome,
                                                                                                    prevtexto_cnpj, prevtexto_codigo, pdf,
                                                                                                    paginainicial, prevpagina)
                            paginas = 0
                        # Se for uma página entra a qui
                        elif paginas == 0:
                            andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                            prevpagina, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo = cria_pdf(sistema, pasta_final, page, matchtexto_nome, matchtexto_cnpj,
                                                                                                    matchtexto_codigo, prevtexto_nome,
                                                                                                    prevtexto_cnpj, prevtexto_codigo, pdf,
                                                                                                    prevpagina, prevpagina)
                except:
                    escreve_relatorio(pasta_final, andamento)
                    continue
                    
            try:
                # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
                if paginas > 0:
                    paginainicial = prevpagina - paginas
                    andamento = '\n' + 'Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1) + '\n\n'
                    cria_pdf(sistema, pasta_final, page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, paginainicial, prevpagina)
                elif paginas == 0:
                    andamento = '\n' + 'Pagina = ' + str(prevpagina + 1) + '\n\n'
                    cria_pdf(sistema, pasta_final, page, matchtexto_nome, matchtexto_cnpj, matchtexto_codigo, prevtexto_nome, prevtexto_cnpj, prevtexto_codigo, pdf, prevpagina, prevpagina)
            except:
                escreve_relatorio(pasta_final, andamento)
                
    separa_duplicatas(pasta_final)
    

if __name__ == '__main__':
    # o robo pega os pdfs na pasta PDFs e cria uma pasta para colocar os separados
    inicio = datetime.now()
    
    separa()
    
    print(datetime.now() - inicio)
    messagebox.showinfo(title=None, message='Separador finalizado')
