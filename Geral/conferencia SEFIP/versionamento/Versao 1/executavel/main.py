# -*- coding: utf-8 -*-
from tkinter import Tk, filedialog, messagebox
import pdfplumber
from socket import socket, AF_INET, SOCK_STREAM
import eel
import os
import re


def porta_em_uso(porta):
    with socket(AF_INET, SOCK_STREAM) as s:
        return s.connect_ex(('localhost', porta)) == 0


eel.init('web', allowed_extensions=['.css', '.html', '.js'])


@eel.expose
def busca_dados():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder = filedialog.askdirectory()
    return folder


def analisa_pdf(arq):
    texto = ''
    with pdfplumber.open(arq) as pdf:
        for pagina in pdf.pages[:]:
            texto += pagina.extract_text() + '\n'

    return texto


def registra_relatorio(caminho, texto=''):
    resumo = os.path.join(caminho, 'resumo.csv')
    if not texto:
        arquivo = open(resumo, 'w')
        arquivo.write('Cod. Cuca;CNPJ;GFIP (Cuca);Diferenca GFIP;INSS (Cuca);Diferenca INSS;Observacao\n')
    else:
        arquivo = open(resumo, 'a')
        arquivo.write(texto+'\n')
    arquivo.close()


def registra_diferenca(caminho, cuca, cnpj, gfip, total, gfip_grf, total_rubrica, obs):
    dif_gfip = float(gfip.replace('.', '').replace(',', '.')) - float(gfip_grf.replace('.', '').replace(',', '.'))
    dif_total = float(total.replace('.', '').replace(',', '.')) - float(total_rubrica.replace('.', '').replace(',', '.'))

    texto = ';'.join([cuca, cnpj, gfip, str(dif_gfip).replace('.', ','), total, str(dif_total).replace('.', ','), obs])
    registra_relatorio(caminho, texto)


def analisa_resumo(resumo):
    print(resumo)
    regex_gfip = re.compile(r'Empresa:\s*(([0-9\.*,]*\s*)+)')
    regex_total = re.compile(r'Valor a Recolher:\s*([0-9\.?]*,\d{2})')


    texto_resumo = analisa_pdf(resumo)
    gfip = regex_gfip.search(texto_resumo)
    total = regex_total.search(texto_resumo)

    total = total.group(1) if total else '0,00'
    try:
        gfip = gfip.group(1)
        gfip = [g.strip('\n') for g in gfip.split(' ') if g != '']
        tamanho = len(gfip)
        gfip = gfip[-3] if tamanho >= 3 else gfip[tamanho-1]
    except:
        gfip = '0,00'

    return total, gfip


def analisa_grf(grf):
    regex_grf = re.compile(r'TOTAL\s*A\s*RECOLHER\s*([0-9\.?]*,\d{2})')

    texto_grf = analisa_pdf(grf)
    gfip_grf = regex_grf.search(texto_grf)
    gfip_grf = gfip_grf.group(1) if gfip_grf else '0,00'
    
    return gfip_grf


def analisa_rubrica(rubrica):
    regex_rubrica = re.compile(r'TOTAL\s*A\s*RECOLHER\s*([0-9\.?]*,\d{2}\s*){4}([0-9\.?]*,\d{2})')

    texto_rubrica = analisa_pdf(rubrica)
    total_rubrica = regex_rubrica.search(texto_rubrica)
    total_rubrica = total_rubrica.group(2) if total_rubrica else '0,00'
    
    return total_rubrica


def get_codigo(arq):
    regex_cod = re.compile(r'(\d*)\s?[A-Z]*')
    texto = analisa_pdf(arq)
    cnpj = texto.split('\n')[2]
    cnpj = ''.join([i for i in cnpj if i.isdigit()])
    codigo = regex_cod.search(texto.split('\n')[1]).group(1)

    return codigo, cnpj


@eel.expose
def iniciar(caminho):
    registra_relatorio(caminho)
    erros = False
    lista_resumo = {}
    lista_grf = {} 
    lista_rubrica = {}
    regex_resumo = re.compile(r'(\d{10,14}) - RESUMO GERAL MENSAL\d?\.[pdfPDF]')
    regex_grf = re.compile(r'^AnaliticoGRF_(\d{10,14})_\d{8}_\d*\.[pdfPDF]')
    regex_rubrica = re.compile(r'^Rubrica_(\d{10,14})_\d{8}_\d*\.[pdfPDF]')

    for arq in os.listdir(caminho):
        extensao = arq[-4:]
        if extensao not in ['.pdf', '.PDF']: 
            continue

        resumo = regex_resumo.search(arq)
        if resumo:
            lista_resumo[resumo.group(1)] = arq

        grf = regex_grf.search(arq)
        if grf:
            n = 11 if len(grf.group(1)) <= 11 else 14
            cnpj = grf.group(1).rjust(n, "0")
            lista_grf[cnpj] = arq

        rubrica= regex_rubrica.search(arq)
        if rubrica:
            n = 11 if len(rubrica.group(1)) <= 11 else 14
            cnpj = rubrica.group(1).rjust(n, "0")
            lista_rubrica[cnpj] = arq

    for cod, arq_resumo in lista_resumo.items():
        obs = ''
        caminho_resumo = os.path.join(caminho, arq_resumo)
        cuca, cnpj = get_codigo(caminho_resumo)
        n = 11 if len(cnpj) <= 11 else 14
        cnpj = cnpj.rjust(n, '0')
        arq_grf = lista_grf.get(cnpj, '')
        arq_rubrica = lista_rubrica.get(cnpj, '')
        if arq_grf or arq_rubrica:
            total, gfip = analisa_resumo(caminho_resumo)
            if arq_grf:
                caminho_gfip = os.path.join(caminho, arq_grf)
                gfip_grf = analisa_grf(caminho_gfip)
            else:
                gfip_grf = '0,00'
                obs = 'AnaliticoGRF nao encontrado'

            if arq_rubrica:
                caminho_rubrica = os.path.join(caminho, arq_rubrica)
                total_rubrica = analisa_rubrica(caminho_rubrica)
            else:
                total_rubrica = '0,00'
                obs = 'Rubrica nao encontrado'

            registra_diferenca(caminho, cuca, cnpj, gfip, total, gfip_grf, total_rubrica, obs)
        else:
            registra_relatorio(caminho, f'{cuca};{cnpj};;;;;Rubrica e AnaliticoGRF nao encontrados')

    return True if not erros else False


porta = 62110
maximo = 62115
while porta <= maximo:
    try:
        if not porta_em_uso(porta):
            eel.start('index.html',  port=porta, size=(600,183))
            break
    except Exception as e:
        print(e)
    porta += 1
else:
    root = Tk()
    root.withdraw()
    messagebox.showwarning("Erro", "Falha ao abrir o programa.")