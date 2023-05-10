# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re
from sys import path
import guias_sinthojur

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice


@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod_sindicato, cnpj, valor = empresa
        _indice(count, total_empresas, empresa)
        
        if cod_sindicato == '39':
            guias_sinthojur.run(cnpj, valor)

if __name__ == '__main__':
    run()