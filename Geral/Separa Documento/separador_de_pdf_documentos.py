# -*- coding: utf-8 -*-
import fitz, re, os, time, shutil
import pyautogui as p
from pathlib import Path
from tkinter.filedialog import askdirectory, Tk
from datetime import datetime
from functools import wraps
from sys import path


def time_execution(func):
    @wraps(func)
    def wrapper():
        comeco = datetime.now()
        print("Execução iniciada as", comeco)
        func()
        print("Tempo de execução", datetime.now() - comeco)

    return wrapper


def ask_for_dir(title='Escolha a pasta onde o robô irá pesquisar'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)

    folder = askdirectory(
        title=title,
    )

    return folder if folder else False


def move_pdf(pdf, file, arq, pasta):
    pdf.close()
    new = os.path.join(pasta, file)
    shutil.move(arq, new)


def cria_pdf(file, page, term, pdf, pasta):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        text = '{} - Página {} da {}'.format(term, str(page.number + 1), str(file))

        # Define o caminho para salvar o pdf
        text = os.path.join(pasta, text)

        # Define a página inicial e a final
        new_pdf.insertPDF(pdf, from_page=page.number, to_page=page.number)

        new_pdf.save(text)


def formata_termo(termo_text):
    termo_text_form = termo_text.replace('$', '\$').replace('.', '\.').replace('[', '\[').replace(']', '\]').replace('?', '\?') \
                                .replace('*', '\*').replace('+', '\+').replace('{', '\{').replace('}', '\}').replace('^', '\^') \
                                .replace('|', '\|').replace('(', '\(').replace('[', '\[')
    return termo_text_form


def escolher_processamento():
    opcao = p.confirm(title='Script incrível', buttons=['Mover resultados para outra pasta', 'Criar uma cópia dos resultados'])
    opcoes = {'Mover resultados para outra pasta': '1', 'Criar uma cópia dos resultados': '2'}
    opcao_escolhida = opcoes[opcao]

    return opcao_escolhida


def separa(termo_text, opcao, ano, doc):
    termo_text_form = formata_termo(termo_text)

    # Definir os padrões de regex
    if termo_text == 'R$ 0,00 R$ 0,00 R$ 0,00 R$ 0,00':
        term = re.compile(r'R\$ 0,00\nR\$ 0,00\nR\$ 0,00\nR\$ 0,00')
    else:
        term = re.compile(r'{}'.format(termo_text_form))

    # data_modificacao = lambda f: f.stat().st_mtime
    res = 0
    arquivos = 0
    for diretorio, subpastas, files in os.walk(doc):
        for subpasta in subpastas:
            '''if subpasta < '17737600000130':
                continue'''
            caminho = f'{doc}/{subpasta}'
            for file in os.listdir(caminho):
                arq = os.path.join(caminho, file)

                # verifica o ano de criação do PDF original
                data_modificacao = time.ctime(os.path.getmtime(arq))
                data_modificacao = data_modificacao.replace('  ', ' ')
                data_modificacao = data_modificacao.split(' ')

                if not file.endswith('.pdf'):
                    continue
                if ano:
                    pasta = os.path.join(f'Resultados {str(data_modificacao[4])}')
                    if data_modificacao[4] != str(ano):
                        continue
                else:
                    pasta = os.path.join(f'Resultados')

                os.makedirs(pasta, exist_ok=True)

                pdf = fitz.open(arq, filetype="pdf")
                print(f'{file}\n{data_modificacao}\n\n')

                # Pra cada pagina do pdf
                for page in pdf:
                    texto = page.getText('text', flags=1 + 2 + 8)
                    # print(f'\n\n---------------------------------------------------------------------------------------------------\n\n{texto}')
                    try:
                        # Procura o termo no texto do pdf
                        match_termo = term.search(texto)
                        if not match_termo:
                            continue

                        # Move o arquivo
                        if opcao == '1':
                            move_pdf(pdf, file, arq, pasta)
                            res += 1
                            break

                        # Copia a página
                        elif opcao == '2':
                            cria_pdf(file, page, termo_text, pdf, pasta)
                            res += 1
                    except:
                        res = 1
                        pdf.close()
                        return res
                if opcao == '2':
                    pdf.close()
                arquivos += 1
        return res, arquivos


@time_execution
def run():
    # pasta final
    termo = p.prompt(title='Script incrível', text='Qual o termo pesquisado?')
    opcao_escolhida = escolher_processamento()
    ano = p.prompt(title='Script incrível', text='Qual o ano de criação do PDF orginal? (Se for irrelevante deixe em branco)')

    # caminho dos PDFs
    documentos = ask_for_dir()

    resultados, arquivos = separa(termo, opcao_escolhida, ano, documentos)

    if int(resultados):
        print(f'{resultados} páginas encontradas com o termo "{termo}" De {arquivos} arquivos pesquisados!')
        p.confirm(title='Script incrível', text=f'Pesquisa finalizada com {resultados} páginas encontradas de {arquivos} arquivos pesquisados!')
    else:
        print(f'Nenhum arquivo de {arquivos} contem o termo pesquisado. "{termo}"')
        p.confirm(title='Script incrível', text=f'Nenhum arquivo de {arquivos} contem o termo pesquisado. "{termo}"')


if __name__ == '__main__':
    run()
