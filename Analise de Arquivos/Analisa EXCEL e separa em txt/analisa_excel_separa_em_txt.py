# -*- coding: utf-8 -*-
import os, re, time
from pyautogui import prompt

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados


def cria_txt(empresa, tipo, encode='latin-1'):
    info1, info2, info3, info4, info5, info6, info7, codigo, cnpj = empresa
    
    local = os.path.join('Execução', str(codigo) + '-' + str(cnpj))
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{tipo} - {cnpj}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{tipo}  - {cnpj} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(';'.join([info1, info2, info3, info4, info5, info6, info7, codigo])) + '\n')
    f.close()


@_time_execution
def run():
    tipo = prompt(title='Script incrível', text='Qual será o nome base do arquivo ".txt"?')
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    for empresa in empresas:
        cria_txt(empresa, tipo)


if __name__ == '__main__':
    run()