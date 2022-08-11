# -*- coding: utf-8 -*-
from datetime import datetime
from Dados import empresas
import fitz, os, re


def separar_gps():
    for arq in os.listdir('arquivos'):
        if not arq.lower().endswith('.pdf'): continue

        file = os.path.join('arquivos', arq)
        pdf = fitz.open(file)
        for i, page in enumerate(pdf):
            texto = page.getText('text', flags=1+2+8)
            dados = re.search(r"3. CÓDIGO DE PAGAMENTO\n([\d\s]{7})\n(\d+)\n", texto)
            if dados:
                pag = dados.group(1).replace(' ', '')
                ident = dados.group(2)
                nome = f"{ident};IMP_DP;15-12-2020;11;2020;Padrão;{pag}.pdf"
                nome = os.path.join('separados', nome)
                with fitz.open() as new_pdf:
                    new_pdf.insertPDF(pdf, from_page=i, to_page=i)
                    new_pdf.save(nome)
            else:
                print(f">>> Falha na leitura da pagina {i}")
        pdf.close()


def separa_dctf():
    regex_name = re.compile(r'F\.Social:\n(.*)\n')
    regex_cnpj = re.compile(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})')

    for  arq in os.listdir(os.path.join('arquivos')):
        if arq[-4:].lower() != '.pdf': continue

        pdf = fitz.open(arq)
        for i, page in enumerate(pdf):
            texto = page.getText('text', flags=1+2+8)

            nome = regex_name.search(texto)
            nome = nome.group(1).replace('/', '').strip() if nome else 'erro'
            cnpj = regex_cnpj.search(texto)
            cnpj  = ''.join(i for i in cnpj.group(1) if i.isdigit()) if cnpj else 'erro'

            if cnpj in empresas: continue

            with fitz.open() as new_pdf:
                new_pdf.insertPDF(pdf, from_page=i, to_page=i)
                new_pdf.save(f'separados/new_way {nome}.pdf')
        pdf.close()


def separa_resumo_geral_mensal():
    caminho = os.path.join('arquivos')

    for arq in os.listdir(caminho):
        if arq[-4:].lower() != '.pdf': continue

        pdf = fitz.open(os.path.join(caminho, arq))
        print(arq)
        old_info = [None] * 2
        for page in pdf:
            texto = page.getText('text', flags=1+2+8) 

            this_info = get_info(texto)
            if old_info[0] == None:
                ipage = fpage = page.number
                old_info = this_info
                continue

            if not this_info or this_info[0] == old_info[0]:
                fpage = page.number
                continue

            with fitz.open() as new_pdf:
                print('arquivo salvo', old_info[1], ipage+1, '-', fpage+1)
                new_pdf.insertPDF(pdf, from_page=ipage, to_page=fpage)
                new_pdf.save(os.path.join('separados', f'{old_info[1]}.pdf'))
                old_info = this_info
                ipage = fpage = fpage + 1
        else:
            with fitz.open() as new_pdf:
                print('arquivo salvo', old_info[1], ipage+1, '-', fpage+1)
                new_pdf.insertPDF(pdf, from_page=ipage, to_page=page.number)
                new_pdf.save(os.path.join('separados', f'{old_info[1]}.pdf'))

        if not pdf.isClosed: pdf.close()
        
    return True


def get_info(texto):
    regex = re.compile(r'\d{5}\n?([A-Z0-9].+)\n')
    regex_cnpj = re.compile(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})')
    regex_razao = re.compile(r'Fax:\n(.*)\n')

    match = regex.search(texto)
    if not match: return False

    cnpj = regex_cnpj.search(match.group(1))
    if cnpj:
        cnpj = ''.join(i for i in cnpj.group(1) if i.isdigit())
        razao = regex_razao.search(texto)
        razao = razao.group(1).replace('/', '')
    else:
        cnpj = regex_cnpj.search(texto)
        cnpj = ''.join(i for i in cnpj.group(1) if i.isdigit())
        razao = match.group(1).replace('/', '')

    return [cnpj, razao]


# Coloque um unico pdf para separação na pasta do script.py e escolha a opção de separação
def run():
    comeco = datetime.now()
    print(comeco)
    os.makedirs('separados', exist_ok=True)

    print('1 - Separa DCTF\n2 - Separa Resumo Geral Mensal\n3 - GPS Individual')
    op = input('Digite a opção: ')

    if op == '1':
        print('\n>>>Separar DCTF')
        separa_dctf()
    elif op == '2':
        print('\n>>>Separar Resumo Geral Mensal')
        separa_resumo_geral_mensal()
    elif op == '3':
        print('\n>>>Separar GPS Individual')
        separar_gps()
    else:
        print("\n>>>Entrada invalida.")

    print("-----" + str(datetime.now()-comeco) + "-----")


if __name__ == '__main__':
    run()



