from datetime import date, timedelta
import pdfplumber
import os, re, sys


def calc_venc(today):
    if today.weekday() == 5 :
        today -= timedelta(days=1)
    elif today.weekday() == 6 :
        today -= timedelta(days=2)
    else:
        pass
    return today.strftime('%d-%m-%Y')


def get_info_gps(tipo, texto):
    regex_cnpj = re.compile(r"IDENTIFICADOR (\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2}|\d{12})")
    regex_comp = re.compile(r"COMPETÊNCIA (\d{2}/\d{4})")

    match = [regex_cnpj.search(texto), regex_comp.search(texto).group(1)]
    if all(match):
        cnpj = match[0].group(1)
        cnpj = "".join(i for i in cnpj if i.isdigit())
        mes, ano = match[1].split('/')
        year, month = date.today().year, date.today().month+1
        venc = calc_venc(date(year, month, 20))
        modelo = f"{cnpj};{venc};IMP_DP;{tipo} - Empresa;{mes};{ano};Padrão;{tipo}.pdf"
        return modelo
    return False


def get_info_fgts(tipo, texto):
    regex = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2}|\d{12}) (\d{2}/\d{4}) (\d{2}/\d{2}/\d{4})")
    match = regex.search(texto)
    
    if match: 
        cnpj = "".join(i for i in match.group(1) if i.isdigit())
        mes, ano = match.group(2).split('/')
        day, month, year = [int(item) for item in match.group(3).split('/')]
        venc = calc_venc(date(year, month, day))
        modelo = f"{cnpj};{venc};IMP_DP;{tipo} - Mensal;{mes};{ano};Padrão;{tipo}.pdf"
        return modelo
    return False


def renomeia_guias_sefip():
    tipos = {
        "FGTS": r"^GRF_\d{11,14}_\d{8}_\d*\.[pdfPDF]{3}$", 
        "GPS": r"^GPS_\d{11,14}_\d{5}_\d{8}_\d*\.[pdfPDF]{3}$"
    }

    for arq in os.listdir('.'):
        for tipo, regex in tipos.items():
            cond = [arq[-4:].lower() == '.pdf', re.search(regex, arq)]

            if all(cond):
                with pdfplumber.open(arq) as pdf:
                    texto = pdf.pages[0].extract_text()
                    modelo = get_info_fgts(tipo, texto) if tipo == "FGTS" else get_info_gps(tipo, texto)     
                if modelo:
                    os.rename(arq, modelo)
                    print(f">>>{tipo} renomeado para {modelo}") 
                    break
                else:
                    print(f">>>{arq} nao pode ser renomeado")
                    break


if __name__ == '__main__':
    renomeia_guias_sefip()
