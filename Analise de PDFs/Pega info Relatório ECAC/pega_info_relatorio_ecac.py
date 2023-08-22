# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def pega_info():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(documentos):
        print(f'\nArquivo: {arq}')
        _escreve_relatorio_csv(f"{arq}")
        # Abrir o pdf
        arq = os.path.join(documentos, arq)
        with fitz.open(arq) as pdf:

            # Para cada página do pdf
            dctf = 'Omissão DCTF'
            for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # print(textinho)
                infos = re.compile(r'Omissão de DCTF').search(textinho)
                cnpj_empresa = re.compile(r'(\d.+)\nUA de Domicílio:\n').search(textinho)
                nome_empresa = re.compile(r'\d\d\d - (.+)\nDados Cadastrais').search(textinho)
                
                if cnpj_empresa:
                    cnpj = cnpj_empresa.group(1)
                    nome = nome_empresa.group(1)
                else:
                    cnpj_empresa = re.compile(r'(\d.+)\nUA de Domicílio: ').search(textinho)
                    if cnpj_empresa:
                        cnpj = cnpj_empresa.group(1)
                        nome = nome_empresa.group(1)
                    
                if infos:
                    try:
                        novo_texto = textinho.split('Diagnóstico Fiscal na Receita Federal ')
                        novo_texto = novo_texto[1]
                    except:
                        novo_texto = textinho
                    
                    periodos = re.findall(r'(\d\d\d\d) - (\w\w.+)', novo_texto)
                    ano_anterior = 0
                    for periodo in periodos:
                        ano = periodo[0]
                        
                        meses = periodo[1]
                        meses = meses.split(' ')
                        
                        competencias = ''
                        for mes in meses:
                            competencias += mes + ';'
                        
                        tipos_dctf = re.findall(r'(Omissão de DCTF)', novo_texto)
                        if len(tipos_dctf) > 1:
                            if int(ano) < ano_anterior:
                                dctf = 'Omissão DCTFWeb*'
                                _escreve_relatorio_csv(f"{cnpj};{nome};{dctf};{ano};{competencias}", nome='Omissão de DCTF')
                            else:
                                _escreve_relatorio_csv(f"{cnpj};{nome};{dctf};{ano};{competencias}", nome='Omissão de DCTF')
                        
                        else:
                            dctf_tipo = re.compile(f'(Omissão de DCT.+) ').search(novo_texto).group(1)
                            _escreve_relatorio_csv(f"{cnpj};{nome};{dctf_tipo};{ano};{competencias}", nome='Omissão de DCTF')
                        ano_anterior = int(ano)
                        
                        if ano_anterior == 0:
                            print(f'\nArquivo: {arq}')
                            print(textinho)
                
                
                
if __name__ == '__main__':
    inicio = datetime.now()
    pega_info()
    print(datetime.now() - inicio)
