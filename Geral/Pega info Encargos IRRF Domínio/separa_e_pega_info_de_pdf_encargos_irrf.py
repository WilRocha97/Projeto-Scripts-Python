# -*- coding: utf-8 -*-
import fitz, re, os, time
from tkinter import messagebox
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from pathlib import Path
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file, ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def guarda_info(page, periodo, nome, cnpj, codigo):
    # guarda info da página anterior
    prev_pagina = page.number
    prev_texto_periodo = periodo
    prev_texto_nome = nome
    prev_texto_cod = codigo
    prev_texto_cnpj = cnpj
    
    return prev_pagina, prev_texto_periodo, prev_texto_nome, prev_texto_cnpj, prev_texto_cod


def cria_pdf(pasta_final, page, periodo, nome, cnpj, codigo, prev_texto_nome, prev_texto_cnpj, pdf, pagina1, pagina2):
    with fitz.open() as new_pdf:
        # Define o nome do arquivo
        nome_pdf = prev_texto_nome.replace('/', ' ').replace(',', ' ')
        cnpj_pdf = prev_texto_cnpj.replace('.', '').replace('/', '').replace('-', '')
        data = periodo.replace('/', '-')
        
        text = f'Encargos IRRF Domínio Web - {cnpj_pdf} - {nome_pdf}.pdf'
        
        # Define o caminho para salvar o pdf
        os.makedirs(os.path.join(pasta_final, f'Relatorios'), exist_ok=True)
        arquivo = os.path.join(pasta_final, f'Relatorios', text)
        
        # Define a página inicial e a final
        new_pdf.insert_pdf(pdf, from_page=pagina1, to_page=pagina2)
        
        new_pdf.save(arquivo)
        
        # atualiza as infos da página anterior
        prev_pagina = page.number
        prev_texto_periodo = periodo
        prev_texto_nome = nome
        prev_texto_cod = codigo
        prev_texto_cnpj = cnpj
        
    return prev_pagina, prev_texto_periodo, prev_texto_nome, prev_texto_cnpj, prev_texto_cod


def separa():
    # Abrir o pdf
    file = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    if not file:
        return False

    pasta_final = 'execução'
    os.makedirs(pasta_final, exist_ok=True)
    
    with fitz.open(file) as pdf:
        prev_pagina = 0
        paginas = 0
        
        # para cada página do pdf
        for page in pdf:
            andamento = f'Pagina = {str(page.number + 1)}'
            '''try:'''
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)

            # Procura o nome da empresa no texto do pdf
            empresa = re.compile(r'(\d+) - (.+)\n(.+)').search(textinho)
            if not empresa:
                continue

            valores = re.compile(r'(\d\d\d\d)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(\d\d/\d\d\d\d)\n(.+)').search(textinho)
            if valores:
                periodo = valores.group(5)
            else:
                periodo = 'Não consta'
                
            codigo = empresa.group(1)
            nome = empresa.group(2)
            cnpj = empresa.group(3)
            
            # Se estiver na primeira página, guarda as informações
            if page.number == 0:
                prev_pagina, prev_texto_periodo, prev_texto_nome, prev_texto_cnpj, prev_texto_cod = guarda_info(page, periodo, nome, cnpj, codigo)
                continue
            
            # Se o nome da página atual for igual ao da anterior, soma um indice de páginas
            if cnpj == prev_texto_cnpj:
                paginas += 1
                # Guarda as informações da página atual
                prev_pagina, prev_texto_periodo, prev_texto_nome, prev_texto_cnpj, prev_texto_cod = guarda_info(page, periodo, nome, cnpj, codigo)
                continue
            
            # Se for diferente ele separa a página
            else:
                if paginas > 0:
                    # define qual é a primeira página e o nome da empresa
                    paginainicial = prev_pagina - paginas
                    andamento = f'Paginas = {str(paginainicial + 1)} até {str(prev_pagina + 1)}'
                    prev_pagina, prev_texto_periodo ,prev_texto_nome, prev_texto_cnpj, prev_texto_cod = cria_pdf(pasta_final, page, periodo, nome, cnpj, codigo,
                                                                         prev_texto_nome, prev_texto_cnpj, pdf,
                                                                         paginainicial, prev_pagina)
                    paginas = 0
                # Se for uma página entra a qui
                elif paginas == 0:
                    andamento = f'Pagina = {str(prev_pagina + 1)}'
                    prev_pagina, prev_texto_periodo, prev_texto_nome, prev_texto_cnpj, prev_texto_cod = cria_pdf(pasta_final, page, periodo, nome, cnpj, codigo,
                                                                         prev_texto_nome, prev_texto_cnpj, pdf,
                                                                         prev_pagina, prev_pagina)
            '''except:
                _escreve_relatorio_csv(andamento, nome='Erro')
                continue'''
        
        # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
        try:
            if paginas > 0:
                paginainicial = prev_pagina - paginas
                andamento = f'Paginas = {str(paginainicial + 1)} até {str(prev_pagina + 1)}'
                cria_pdf(pasta_final, page, periodo, nome, cnpj, codigo, prev_texto_nome, prev_texto_cnpj, pdf, paginainicial,
                         prev_pagina)
            elif paginas == 0:
                andamento = f'Pagina = {str(prev_pagina + 1)}'
                cria_pdf(pasta_final, page, periodo, nome, cnpj, codigo, prev_texto_nome, prev_texto_cnpj, pdf, prev_pagina,
                         prev_pagina)
        except:
            _escreve_relatorio_csv(andamento, nome='Erro')
    
    return file


def analiza(file):
    # Abrir o pdf
    sem_codigo_receita = 'não'
    with fitz.open(file) as pdf:
        
        # Para cada página do pdf, se for a segunda página o script ignora
        for page in pdf:
            
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                # Procura a descrição do valor a recolher
                valores = re.compile(r'(\d\d\d\d)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(\d\d/\d\d\d\d)\n(.+)').findall(textinho)
                
                # se não encontrar valores com o código da receita procura valor de IRRF
                if not valores:
                    # se existir uma empresa diferente sem código anterior a empresa atual, anota na planilha
                    if sem_codigo_receita == 'sim':
                        if codigo != codigo_anterior:
                            _escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_anterior};Não consta;{periodo_anterior};Não consta;{valor_anterior};Não consta;Não consta;Não consta")
    
                    try:
                        valor_anterior = re.compile(r'Total Geral:\n(.+,\d+)').search(textinho).group(1)
                        try:
                            periodo_anterior = re.compile(r'Mensal (\d\d/\d\d)').search(textinho).group(1)
                        except:
                            periodo_anterior = 'Não consta'
                        
                        # armazena os valores para compara comparar com a empresa da próxima página
                        empresa = re.compile(r'(\d+) - (.+)\n(.+)').search(textinho)
                        codigo_anterior = empresa.group(1)
                        nome_anterior = empresa.group(2)
                        cnpj_anterior = empresa.group(3)
                        sem_codigo_receita = 'sim'
                    except:
                        pass
                else:
                    for valor in valores:
                        empresa = re.compile(r'(\d+) - (.+)\n(.+)').search(textinho)
                        codigo = empresa.group(1)
                        nome = empresa.group(2)
                        cnpj = empresa.group(3)
                        
                        # se existir uma empresa diferente sem código anterior a empresa atual, anota na planilha
                        if sem_codigo_receita == 'sim':
                            if codigo != codigo_anterior:
                                _escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_anterior};Não consta;{periodo_anterior};Não consta;{valor_anterior};Não consta;Não consta;Não consta")
                                sem_codigo_receita = 'não'
                                
                        periodicidade = valor[6]
                        periodo = valor[5]
                        codigo_receita = valor[0]
                        valor_recolher = valor[1]
                        valor_compensar = valor[2]
                        valor_pagar = valor[3]
                        valor_acumular = valor[4]
                        
                        # print(f'{codigo} - {cnpj} - {nome} - {valor}')
                        
                        _escreve_relatorio_csv(f"{codigo};{cnpj};{nome};{periodicidade};{periodo};{codigo_receita};{valor_recolher};{valor_compensar};{valor_pagar};{valor_acumular}")
                        sem_codigo_receita = 'não'
                            
            except():
                print(textinho)
                print('ERRO')
    
    return periodo


if __name__ == '__main__':
    file = separa()

    inicio = datetime.now()
    periodo = analiza(file)
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;PERIODICIDADE;PERÍODO;CÓDIGO RECOLHIMENTO;VALOR A RECOLHER;VALOR A COMPENSAR;VALOR A PAGAR;VALOR A ACUMULAR')
    print(datetime.now() - inicio)
    
    messagebox.showinfo(title=None, message='Finalizado com sucesso')
