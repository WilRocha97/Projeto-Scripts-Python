# -*- coding: utf-8 -*-
import os, csv

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers


def gera():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    for count1, empresa in enumerate(empresas, start=1):
        cnpj, nome, valor, cod = empresa
        cnpj_armazenado = cnpj
        nome_armazenado = nome
        cod_armazenado = cod
        valor_armazenado = float(valor.replace(',', '.'))
        
        for count2, empresa in enumerate(empresas, start=1):
            cnpj, nome, valor, cod = empresa
            if count2 != count1:
                if cnpj_armazenado == cnpj and cod_armazenado == cod:
                    valor_armazenado += float(valor.replace(',', '.'))

        sp = str(valor_armazenado).split('.')
        valor_armazenado = f'{str(sp[0])},{str(sp[1][:2])}'
        
        _escreve_relatorio_csv(f'{cnpj_armazenado};{nome_armazenado};{valor_armazenado};{cod_armazenado}', nome='Dados Somados')


def tira_duplicadas():
    with open('execucao/Dados Somados.csv', encoding='utf-8') as arquivo_referencia:
        # 2. ler a tabela
        tabela = csv.reader(arquivo_referencia, delimiter=';')

        linha_armazenada = ''
        for linha in tabela:
            if linha != linha_armazenada:
                _escreve_relatorio_csv(f'{linha[0]};{linha[1]};{linha[2]};{linha[3]}', nome='Dados Somados Filtrados')
            linha_armazenada = linha
                    
                
@_time_execution
def run():
    os.makedirs('execucao', exist_ok=True)
    # p.mouseInfo()
    gera()
    tira_duplicadas()
    

if __name__ == '__main__':
    run()
