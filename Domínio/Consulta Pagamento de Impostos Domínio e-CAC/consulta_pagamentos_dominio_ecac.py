# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login


def relatorio_pagamento_ecac(empresa, periodo, andamento):
    cod, cnpj, nome = empresa
    _wait_img('Movimentos.png', conf=0.9, timeout=-1)
    
    p.hotkey('alt', 'm')
    time.sleep(1)
    p.press('g')
    time.sleep(1)
    p.press('t')

    _wait_img('ConsultarPagamentos.png', conf=0.9, timeout=-1)
    
    p.write(periodo)
    
    if _find_img('ComCertDesmarcado.png', conf=0.9):
        _click_img('ComCertDesmarcado.png', conf=0.9)
    time.sleep(1)
    
    p.hotkey('alt', 'c')
    
    while not _find_img('INSIRA AQUI.png', conf=0.9):
        time.sleep(1)
        if _find_img('NaoAcessouEcac.png', conf=0.9):
            print('❌ Não é possível acessar o e-CAC, tentando novamente.')
            p.press('enter')
            _wait_img('ConsultarPagamentos.png', conf=0.9, timeout=-1)
            time.sleep(1)
            p.hotkey('alt', 'c')
            
        if _find_img('NaoAcessouEcac.png', conf=0.9):
            print('❌ Certificado Digital não encontrado, tentando novamente.')
            p.press('enter')
            _wait_img('ConsultarPagamentos.png', conf=0.9, timeout=-1)
            time.sleep(1)
            p.hotkey('alt', 'c')
    
    p.press('esc', presses=3)


@_time_execution
def run():
    periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Relatórios para DARF DCTF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
    
        if not _login(empresa, andamentos):
            continue
        relatorio_pagamento_ecac(empresa, periodo, andamentos)


if __name__ == '__main__':
    run()
