# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login, _salvar_pdf


def arquivos_dirf(empresa, ano, andamento):
    cod, cnpj, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    
    # under construction
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Arquivo gerado']), nome=andamento)
    print('✔ Arquivo gerado')

    p.press('esc', presses=4)
    time.sleep(2)


@_time_execution
def run():
    ano = p.prompt(text='Qual ano base?', title='Script incrível', default='0000')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos DIRF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
    
        if not _login(empresa, andamentos):
            continue
        arquivos_dirf(empresa, ano, andamentos)


if __name__ == '__main__':
    run()
