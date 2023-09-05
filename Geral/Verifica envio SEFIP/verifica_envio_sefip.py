# -*- coding: utf-8 -*-
from pandas import read_excel, read_csv
from datetime import datetime
import pdfplumber
import shutil 
import os
import re


#_CAMINHO = os.path.join(r'\\', 'srv001', 'Sefip_Guias', 'RELATORIOS_SEFIP', 'SEFIP_RELATÓRIOS_2020', '08_2020')
_CAMINHO = os.path.join('08_2020')


def get_codigo(rubrica, lista_rubrica, lista_sefip):
    regex = re.compile(r'Nº ARQUIVO: (.+-\d)')
    lista = []
    for arq in lista_rubrica:
        with pdfplumber.open(os.path.join(rubrica, arq)) as pdf:
            texto = pdf.pages[0].extract_text()
            if not texto: continue
            numero = regex.search(texto)
            if numero:
                codigo = numero.group(1).replace('-','')
                print(arq, '->', codigo)
                enviar = 'Enviada' if codigo in lista_sefip else 'Não Enviada'
                lista.append((arq, codigo, enviar))
            else:
                print('>>> ERRO NO NUMERO NO ARQUIVO ', arq)
    return lista
    

def mover_arquivos(rubrica, sefip):
    regex_rubrica = re.compile(r'Rubrica_(\d)+_(.+)\.pdf', re.I)
    regex_sefip = re.compile(r'SEFIP\s*(.+)\.pdf', re.I)
    lista_rubrica = []
    lista_sefip = []
    for raiz, subpastas, arquivos in os.walk(_CAMINHO):
        for arquivo in arquivos:
            nome_sefip = regex_sefip.search(arquivo)
            new = ''
            if regex_rubrica.search(arquivo):
                if arquivo not in lista_rubrica:
                    lista_rubrica.append(arquivo)
                    new = os.path.join(rubrica, arquivo)
            elif nome_sefip:
                if nome_sefip.group(1) not in lista_sefip:
                    lista_sefip.append(nome_sefip.group(1))
                    new = os.path.join(sefip, arquivo)
            if new and raiz not in ['rubrica', 'sefip']:
                old = os.path.join(raiz, arquivo)
                shutil.move(old, new)

    return lista_rubrica, lista_sefip


def analisa_dpcuca():
    colunas = ['cod', 'cnpj', 'razao']
    arquivo = read_csv('dpcuca.csv', header=None, names=colunas, sep=';', encoding='latin-1')
    lista = []

    for linha in arquivo.itertuples():
        conds = [linha.cod, linha.cnpj, linha.razao]
        if all(conds):
            codigo = str(conds[0])
            cpf_cnpj = "".join([i for i in conds[1] if i.isdigit()])
            lista.append([codigo, str(conds[2]), cpf_cnpj]) 

    return lista


def analisa_plan_neuci():
    colunas = [
        'cod', 'cliente', 'prot', 'folha', 'c_905', 'c_115', 'IRF', 'PIS', 
        'sind', 'pensao', 'ret_11', 'sindical', 'RPA', 'prolabore', '',
    ]
    lista = []
    arquivo = read_excel('RELAÇÃO EMPRESAS.XLS', names=colunas, keep_default_na=False, header=1)
    filtro = arquivo.query("c_905=='X' | folha=='X' | c_115=='X'")
    for linha in filtro.itertuples():
        if 900 <= int(linha.cod) <= 999: break
        lista.append(str(linha.cod))

    return lista


def confere_rubrica(lista):
    relatorio = open('relatorio.csv', 'w')
    relatorio.write("codigo;cpf_cnpj;protocolo;situação rubrica;situacao sefip\n")
    lista_dp = analisa_dpcuca()
    lista_neuci = analisa_plan_neuci()
    lista_empresas = [(empresa[2], codigo) for codigo in lista_neuci for empresa in lista_dp if codigo == empresa[0]]
    
    for empresa, cod in lista_empresas:
        texto = ''
        for item in lista:
            cnpj = item[0].split('_')[1]
            if cnpj in empresa:
                texto += ";".join([cod, empresa, item[1], item[0], item[2]])+"\n"
        if not texto:
            texto = ";".join([cod, empresa,"-","Não Encontrado","Não Enviado"])+"\n"
        relatorio.write(texto)

    relatorio.close()



if __name__ == '__main__':
    comeco = datetime.now()
    rubrica = os.path.join(_CAMINHO, 'rubrica')
    sefip = os.path.join(_CAMINHO, 'sefip')
    os.makedirs(rubrica, exist_ok=True)
    os.makedirs(sefip, exist_ok=True)

    print('>>> Movendo arquivos')
    lista_rubrica, lista_sefip = mover_arquivos(rubrica, sefip)

    print('>>> Analisando Rublicas')
    lista = get_codigo(rubrica, lista_rubrica, lista_sefip)

    print('>>> Conferindo a lista')
    confere_rubrica(lista)

    print('>>> Programa finalizado em', (datetime.now()-comeco))
