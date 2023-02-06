# -*- coding: utf-8 -*-
import os, csv, decimal
from decimal import Decimal
from pathlib import Path

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers

e_dir = Path('ignore')
ctx = decimal.getcontext()
ctx.rounding = decimal.ROUND_HALF_UP


def gera():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    for count1, empresa in enumerate(empresas, start=1):
        cnpj, nome, nota, valor, cod = empresa
        valor = Decimal(float(valor.replace(',', '.')))
        cnpj_armazenado = cnpj
        nome_armazenado = nome.replace('/', ' ')
        nota_armazenado = [nota]
        cod_armazenado = cod
        valor_armazenado = valor
        
        for count2, empresa in enumerate(empresas, start=1):
            cnpj, nome, nota, valor, cod = empresa
            valor = Decimal(float(valor.replace(',', '.')))
            
            if count2 != count1:
                if cnpj_armazenado == cnpj and cod_armazenado == cod:
                    valor_armazenado += valor
                    nota_armazenado.append(nota)
                    
        valor_armazenado = round(valor_armazenado, 2)
        valor_armazenado = str(valor_armazenado).replace('.', ',')
        
        # organiza os números das notas em orden crescente
        nota_armazenado.sort()
        notas_armazenadas = ''
        for nota in nota_armazenado:
            notas_armazenadas += f'{nota}, '

        notas_armazenadas += ', '
        _escreve_relatorio_csv(f'{cnpj_armazenado};{nome_armazenado};{notas_armazenadas.replace(", , ", "")};{valor_armazenado};{cod_armazenado}', nome='Dados Somados')


def tira_duplicadas():
    with open('execução/Dados Somados.csv', encoding='latin-1') as arquivo_referencia:
        # 2. ler a tabela
        tabela = csv.reader(arquivo_referencia, delimiter=';')
        
        linha_armazenada = ['', '', '', '', '']
        for linha in tabela:
            if linha != linha_armazenada:
                _escreve_relatorio_csv(f'{linha[0]};{linha[1]};{linha[2]};{linha[3]};{linha[4]}', nome='Dados', local=e_dir)
            linha_armazenada = linha
                    
                
@_time_execution
def run():
    os.makedirs('execução', exist_ok=True)
    # p.mouseInfo()
    gera()
    tira_duplicadas()
    os.remove('execução/Dados Somados.csv')
    

if __name__ == '__main__':
    run()
