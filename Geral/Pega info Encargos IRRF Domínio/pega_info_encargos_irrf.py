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
                            _escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_anterior};Não consta;{periodo_anterior};Não consta;{valor_anterior};Não consta;Não consta;Não consta", nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
                
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
                                _escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_anterior};Não consta;{periodo_anterior};Não consta;{valor_anterior};Não consta;Não consta;Não consta", nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
                                sem_codigo_receita = 'não'
                    
                        periodicidade = valor[6]
                        periodo = valor[5]
                        codigo_receita = valor[0]
                        valor_recolher = valor[1]
                        valor_compensar = valor[2]
                        valor_pagar = valor[3]
                        valor_acumular = valor[4]
                    
                        # print(f'{codigo} - {cnpj} - {nome} - {valor}')
                    
                        _escreve_relatorio_csv(f"{codigo};{cnpj};{nome};{periodicidade};{periodo};{codigo_receita};{valor_recolher};{valor_compensar};{valor_pagar};{valor_acumular}", nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
                        sem_codigo_receita = 'não'
        
            except():
                print(textinho)
                print('ERRO')

    return periodo


if __name__ == '__main__':
    inicio = datetime.now()
    periodo = analiza()
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;PERIODICIDADE;PERÍODO;CÓDIGO RECOLHIMENTO;VALOR A RECOLHER;VALOR A COMPENSAR;VALOR A PAGAR;VALOR A ACUMULAR', nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
    print(datetime.now() - inicio)
