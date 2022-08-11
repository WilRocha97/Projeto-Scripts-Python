# -*- coding: utf-8 -*-
from pandas import read_excel, read_csv
import re


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
    arquivo = read_excel('RELAÇÃO EMPRESAS.XLS', names=colunas, keep_default_na=False, header=1)
    filtro = arquivo.query("c_905=='X'")
    lista = [str(x.cod) for x in filtro.itertuples()]

    return lista


# GERA UMA PLANILHA CRUZANDO OS DADOS DO DPCUCA COM AS EMPRESAS LISTADAS
# NO ARQUIVO DA NEUCI QUE POSSUEM O CÓDIGO 905 MARCADO PARA EMISSÃO DE PROLABORE
# UTILIZADO PARA EMISSÃO DOS RELATORIOS "RESUMO GERAL MENSAL"
if __name__ == '__main__':
    lista_dp = analisa_dpcuca()
    lista_neuci = analisa_plan_neuci()
    lista_empresas = [empresa for codigo in lista_neuci for empresa in lista_dp if codigo == empresa[0]]
    
    arquivo = open('lista_empresas.csv', 'w')
    for empresa in lista_empresas:
        arquivo.write(';'.join(empresa)+'\n')
    arquivo.close()
