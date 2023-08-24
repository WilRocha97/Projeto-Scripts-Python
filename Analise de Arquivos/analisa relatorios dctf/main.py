from datetime import datetime
from pandas import read_csv, DataFrame
from shutil import rmtree
from time import sleep
import fitz, re, os


def escreve_relatorio_csv(texto, end='\n'):
    try:
        f = open('Resumo.csv', 'a')
    except:
        f = open('Resumo - erro.csv', 'a')
    f.write(texto + end)
    f.close()


def search_in_relatorio():
    regex = re.compile(r'(.*)\n(\d{2}\/\d{2}\/\d{4})\n(\d{2})\n\s+([\d\.]+,\d{2})\n(\d+)')

    pdf = fitz.open(os.path.join('dados', 'relatorios.pdf'))
    for page in pdf:
        control = False
        texto = page.getText('text', flags=1+2+8)

        match = re.search(r"F\.Social:\n(.*)\n.*\n([\d\.\-/]{18})\nCNPJ:", texto)
        if not match: continue

        nome, cnpj = match.groups()
        for match in regex.finditer(texto):
            tipo, venc, dig, valor, cod = match.groups()
            tipo = tipo.strip()
            cod = f"{cod}-{str(dig).rjust(2, '0')}"

            escreve_relatorio_csv(f"{nome};{cnpj};{tipo};{cod};{venc};{valor}")
            search_in_darf(nome, cnpj)
            control = True

        if not control:
            escreve_relatorio_csv(f"{nome};{cnpj};Sem debitos no relátorio")


def search_in_darf(nome, cnpj):
    results = []
    cod = re.compile(r"Receitas Federais\n?([\d\s]{7})")
    venc = re.compile(r"06 Data de Vencimento\n?([\d\/]{10})")
    valor = re.compile(r"10 Valor Total\n?\s+([\d\.]+,\d{2})")

    darfs = os.path.join('temp', 'darfs', cnpj)
    if not os.path.exists(darfs): return False

    for darf in os.listdir(darfs):
        pdf = fitz.open(os.path.join(darfs, darf))
        for page in pdf:
            texto = page.getText('text', flags=1+2+8)
            info = [cod.search(texto), venc.search(texto), valor.search(texto)]
            if not all(info): continue

            info[0] = ''.join(i for i in info[0].group(1) if i.isdigit())
            info[1] = info[1].group(1)
            info[2] = info[2].group(1)

            texto = f"{nome};{cnpj};Extraido da guia;" + ";".join(info)
            escreve_relatorio_csv(texto)
            break

        pdf.close()

    return True


def separador(tipo, arq):
    print('>>>separando', tipo)
    pasta = os.path.join('temp', tipo)
    os.makedirs(pasta, exist_ok=True)
    cnpj_regex = re.compile(r"([\d\.\/-]{14,18})\n")

    pdf = fitz.open(arq)
    for page in pdf:
        texto = page.getText('text', flags=1+2+8)

        cnpj = cnpj_regex.search(texto)
        if not cnpj: continue
        cnpj = ''.join(i for i in cnpj.group(1) if i.isdigit())

        os.makedirs(os.path.join(pasta, cnpj), exist_ok=True)
        with fitz.open() as new_pdf:
            new_pdf.insertPDF(pdf, from_page=page.number, to_page=page.number)
            new_name = os.path.join(pasta, cnpj, f'{page.number}.pdf')
            new_pdf.save(new_name)
    pdf.close()


def run():
    inicio = datetime.now()

    _relatorio = os.path.join('dados', 'relatorios.pdf')
    _darfs = os.path.join('dados', 'darfs.pdf')

    try:
        rmtree('temp')
    except:
        pass
    
    separador('darfs', _darfs)
    escreve_relatorio_csv('Razao;CNPJ;TIPO TRIBUTO;COD. RECEITA;VENCIMENTO;VALOR')

    print('>>>percorrendo lista')
    search_in_relatorio()

    rmtree('temp')
    print(datetime.now()-inicio)
    return True


if __name__ == '__main__':
    run()