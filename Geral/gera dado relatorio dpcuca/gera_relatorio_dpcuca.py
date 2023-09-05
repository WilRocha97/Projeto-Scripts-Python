# -*- coding: utf-8 -*-
from Dados import lista
from datetime import datetime
import datedelta


def verifica(mes):
    if mes < 10:
        return 'D' + str(mes)
    elif mes == 10:
        return 'A'
    elif mes == 11:
        return 'B'
    else:
        return 'C'
        
def gera_relatorio_dpcuca(cod, nome, cnpj, inicio, fim):
    inicio = [int(i) for i in inicio.split('-')]
    fim = [int(i) for i in fim.split('-')]

    dt_inicio = datetime(inicio[1], inicio[0], 1, 0, 0)
    dt_fim = datetime(fim[1], fim[0], 1, 0, 0)

    while dt_inicio <= dt_fim:
        #index = verifica(dt_inicio.month)
        #print(f'{cod};{nome};{index};{dt_inicio.year};{dt_inicio.month}')
        print(f'{cod};{"".join(i for i in cnpj if i.isdigit())};{cnpj};{nome.upper()};{dt_inicio.month};{dt_inicio.year}')
        dt_inicio += datedelta.MONTH


def run():
    for linha in lista:
        gera_relatorio_dpcuca(*linha)

if __name__ == '__main__':
    run()