# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from pyperclip import copy
import os

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def excluir_documento():
    while not _find_img('abrir_arquivo.png', conf=0.9):
        time.sleep(1)
        p.press('up')
        _click_img('minimizar.png', conf=0.9)
    
    time.sleep(1)
    p.hotkey('alt', 'e')
    time.sleep(1)
    p.hotkey('alt', 's')
    
    # espera exlcuir o arquivo
    _wait_img('arquivo_excluido.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    
@_time_execution
def run():
    while 1 > 0:
        excluir_documento()


if __name__ == '__main__':
    run()
