# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def analiza():
    arq = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    # Abrir o pdf
    with fitz.open(arq) as pdf:
        # Para cada página do pdf, se for a segunda página o script ignora
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                """if page.number == 3:
                    print(textinho)
                    time.sleep(33)"""
                dados = re.compile(r'(.+)\n(.+)\n(.+)\n(.+)\n.+\nAté:').findall(textinho)
                
                for dado in dados:
                    nome = dado[3]
                    cpf = dado[0]
                    cpf = cpf.replace('.', '').replace('-', '')
                    try:
                        int(cpf)
                        if cpf == '0':
                            cpf = re.compile(r'(.+)\n(.+)\n' + nome).search(textinho)
                            cpf = cpf.group(1)
                            cpf = cpf.replace('.', '').replace('-', '')
                            try:
                                int(cpf)
                            except:
                                cpf = 'CPF não encontrado'
                    except:
                        try:
                            cpf = re.compile(r'(.+)\n(.+)\n' + nome).search(textinho)
                            cpf = cpf.group(1)
                            cpf = cpf.replace('.', '').replace('-', '')
                            int(cpf)
                        except:
                            cpf = 'CPF não encontrado'
                    
                    empregado_cod = ''
                    empregado_nome_separado = ''
                    for i in range(150):
                        empregado = re.compile(r'(.+)\n(.+)\n(.+\n){' + str(i) + '}' + nome).search(textinho)
                        if empregado.group(1) == 'Empregado:':
                            empregado_nome = empregado.group(2)
                            
                            for i in str(empregado_nome):
                                try:
                                    int(i)
                                    empregado_cod += str(i)
                                except:
                                    empregado_nome_separado += str(i)
                            break
                    
                    if empregado_nome_separado == '':
                        empregado_nome_separado = re.compile(r'(.+)\nCódigo:').search(textinho).group(1)
                    
                    try:
                        id_empregador = re.compile(r'(.+)\nCNPJ').search(textinho).group(1)
                    except:
                        try:
                            id_empregador = re.compile(r'(.+)\nCAEPF').search(textinho).group(1)
                        except:
                            id_empregador = re.compile(r'(.+)\nCPF').search(textinho).group(1)
                            
                    nome_empregador = re.compile(r'Empresa:\n(.+)').search(textinho).group(1)
                    
                    _escreve_relatorio_csv(f'{id_empregador};{nome_empregador};{empregado_cod};{empregado_nome_separado};{nome};{cpf}', nome='Dependentes')
                    print(cpf)
                
            except():
                print(textinho)
                print('ERRO')
            

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ | CPF;NOME EMPRESA;CÓD EMPREGADO;EMPREGADO;DEPENDENTE;CPF', nome='Dependentes')
    print(datetime.now() - inicio)
