# -*- coding: utf-8 -*-
from os import makedirs, path, listdir
from tkinter import Tk, filedialog
from shutil import rmtree
from time import sleep
import subprocess, platform, fitz, eel, re


eel.init('web', allowed_extensions=['.css', '.html', '.js'])

@eel.expose
def busca_dados():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder = filedialog.askopenfilename(filetypes=[("PDF file", ".pdf"), ])
    
    root.destroy()
    return folder


def compress_pdf(arqs, tipo):
    makedirs(f'{tipo}_compress', exist_ok=True)

    if platform.machine().endswith('64'):
        gs_path = 'C:\\Program Files\\gs\\gs9.55.0\\bin\\gswin64c.exe'
    else:
        gs_path = 'C:\\Program Files (x86)\\gs\\gs9.55.0\\bin\\gswin32c.exe'

    for file in arqs:
        filename = path.join(f'{tipo}_temp', file)
        arg1 = '-sOutputFile=' + path.join(f'{tipo}_compress', file)
        p = subprocess.Popen(args=[
            str(gs_path), '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
            '-dPDFSETTINGS=/screen', '-dNOPAUSE', '-dBATCH', '-dQUIET', str(arg1), filename
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)

    return True


def separa_pdf(arq, tipo):
    regexs = [re.compile(r"([\d\./\-]{14,18})\n")]
    if tipo == 'darf':
        regexs.extend([
            re.compile(r"06 Data de Vencimento\n(\d{2}/\d{2}/\d{4})\n"),
            re.compile(r"02 Período de Apuração\s+\n\d{2}/(\d{2}/\d{4})\n"),
            re.compile(r"Receitas Federais\n([\d\s]{7})\n")
        ])
    elif tipo == 'gps':
        regexs.extend([
            re.compile(r"(\d{11})\n4\. COMPETÊNCIA"),
            re.compile(r"INSS\n(\d{4})(?:.|\n)+G P S\n/\n(\d{2})"),
            re.compile(r"3\. CÓDIGO DE PAGAMENTO\n([\d\s]{7})\n"),
        ])

    pdf = fitz.open(arq)
    paginas = []

    for page in pdf:
        texto = page.getText('text', flags=1+2+8)

        results = []
        for regex in regexs:
            aux = regex.search(texto)
            if not aux:
                paginas.append(str(page.number))
                continue
            results.append("".join(i for i in "".join(aux.groups()) if i.isdigit()))

        makedirs(f'{tipo}_temp', exist_ok=True)
        nome = path.join(f'{tipo}_temp', ';'.join(results) + f';{tipo.upper()}.pdf')

        with fitz.open() as new_pdf:
            new_pdf.insertPDF(pdf, from_page=page.number, to_page=page.number)
            new_pdf.save(nome)

    list_arq = listdir(f'{tipo}_temp')
    if list_arq: compress_pdf(list_arq, tipo)
    sleep(2)

    pdf.close()
    rmtree(f"{tipo}_temp")
    return ', '.join(paginas)


@eel.expose
def iniciar(file, tipo):
    try:
        paginas = separa_pdf(file, tipo)
        return paginas or True
    except Exception as e:
        print(e)
        return False


eel.start('index.html',  port=0, size=(600,230))