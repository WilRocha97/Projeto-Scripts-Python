# -*- coding: utf-8 -*-
import datetime, pyperclip, time, os, shutil, pyautogui as p
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login, _salvar_pdf


def define_competencia(mes):
    return data_inicial, data_final


def dirf(empresa, andamento, data_inicial, data_final):
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Encargos
    p.press('e')
    time.sleep(0.5)
    # IRRF
    time.sleep(0.5)
    p.press('i')
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)
    
    # digita o ano base
    p.write(ano)
    time.sleep(0.5)
    p.press('tab')
    
    # digita o código do responsável
    p.write('1')
    time.sleep(0.5)
    p.press('tab')
    
    p.press('esc', presses=4)
    time.sleep(2)


@_time_execution
def run():
    mes = p.prompt(text='Qual a competencia?', title='Script incrível', default='00')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos DIRF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    data_inicial, data_final = define_competencia(mes)
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
    
        if not _login(empresa, andamentos):
            continue
        irrf(empresa, andamentos, data_inicial, data_final)


if __name__ == '__main__':
    run()
