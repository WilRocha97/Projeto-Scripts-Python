# -*- coding: utf-8 -*-
import time, fitz, re, os
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime

# Definir os padrões de regex
padrao_valor = re.compile(r'Valor a Recolher.\n.+ (.+)')
padrao_nome = re.compile(r'Resumo por Eventos do Mês\n.(\d{1,4})(\w.+).\n.+.\n.+.\n.+.\n(.+)')
padrao_ref1 = re.compile(r'Fax:\n (.+)')
padrao_ref2 = re.compile(r'Fax:.+\n (.+)')
padrao_ref3 = re.compile(r'Fax:\n.+\n(.+)')
padrao_ref4 = re.compile(r'Fax:.+\n.+\n(.+)')


def ask_for_dir(title=''):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(title=title)
    
    return folder if folder else False


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Resumo Mensal.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Resumo Mensal - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')


def analiza():
    documentos = ask_for_dir(title='Selecione a pasta onde estão os PDFs para separar')
    if not documentos:
        return False
    
    # Analiza cada pdf que estiver na pasta
    for file in os.listdir(documentos):
        print('\nArquivo: {} \n'.format(file))
        # Abrir o pdf
        with fitz.open(r'PDF\{}'.format(file)) as pdf:

            # Para cada página do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # Procura o valor a recolher da empresa
                    localiza_valor = padrao_valor.search(textinho)

                    if not localiza_valor:
                        continue

                    # Procura o nome da empresa
                    localiza_nome = padrao_nome.search(textinho)
                    # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a éssa info
                    localiza_ref = padrao_ref1.search(textinho)

                    if not localiza_ref:
                        # Procura a descrição do valor a recolher 2
                        localiza_ref = padrao_ref2.search(textinho)
                        if not localiza_ref:
                            # Procura a descrição do valor a recolher 3
                            localiza_ref = padrao_ref3.search(textinho)
                            if not localiza_ref:
                                # Procura a descrição do valor a recolher 4
                                localiza_ref = padrao_ref4.search(textinho)

                    # Guarda as infos da empresa
                    cod = localiza_nome.group(1)
                    nome = localiza_nome.group(2)
                    cnpj = localiza_nome.group(3)
                    valor = localiza_valor.group(1)
                    ref = localiza_ref.group(1)

                    # print('{} - {} - R${} - {}\n'.format(cod, cnpj, valor, nome))

                    texto = ';'.join([cod, cnpj, nome, ref, valor])
                    escrever(texto)

                except():
                    print('ERRO')
                    print(textinho)


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
    messagebox.showinfo(title=None, message='Finalizado')
    