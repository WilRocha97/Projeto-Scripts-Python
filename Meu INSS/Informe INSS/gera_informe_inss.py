import pyautogui as p, PySimpleGUI as sg
import time
import os
import pyperclip
from threading import Thread

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def login(empresa, andamentos):
    
    time.sleep(22)
    
    return resultado

@_time_execution
@_barra_de_status
def run(window):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)

        cnpj, comp, ano, venc = empresa
        andamentos = f'Informes INSS'

        _abrir_chrome('https://meu.inss.gov.br/#/login')
        resultado = login(empresa, andamentos)



        # Salvar a guia
        imprimir(empresa, andamentos)

    time.sleep(2)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
