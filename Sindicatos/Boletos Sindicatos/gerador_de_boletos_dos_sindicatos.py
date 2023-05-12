# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re
from sys import path
import guias_sinthojur, guias_sindpd

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
        cod_sindicato, cnpj, valor, usuario, senha, funcionarios = empresa
        _indice(count, total_empresas, empresa)
        
        sindicatos = {
            '3': '',
            '8': '',
            '10': '',
            '11': '',
            '16': '',
            '17': '',
            '19': '',
            '21': '',
            '22': '',
            '23': '',
            '25': '',
            '28': guias_sindpd.run,
            '39': guias_sinthojur.run,
            '49': '',
            '58': '',
            '65': '',
            '69': '',
            '70': '',
            '100': '',
            '131': '',
            '133': '',
            '135': '',
            '148': '',
            '162': '',
            '223': ''
        }
        
        resultado = sindicatos[cod_sindicato](cnpj, valor, usuario, senha, funcionarios)
        _escreve_relatorio_csv(f'{cnpj};{valor};{resultado}', nome='Boletos Sindicato')
        print(resultado)
        
            
if __name__ == '__main__':
    run()