# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login, _salvar_pdf


def relatorio_darf_dctf(empresa, andamento):
    cod, cnpj, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Acompanhamentos
    p.press('a')
    # Demonstrativo Mensal
    time.sleep(0.5)
    p.press('m')
    
    while not _find_img('demonstrativo_mensal.png', conf=0.9):
        time.sleep(1)
    
    # configura o ano e digita no domínio
    ano = datetime.year
    
    time.sleep(1)
    p.write(f'01{ano}')
    
    time.sleep(0.5)
    p.press('tab')
    
    time.sleep(1)
    p.write(f'12{ano}')
    
    # gera o relatório
    time.sleep(1)
    p.press('alt', 'o')
    
    # espera gerar
    while not _find_img('demonstrativo_mensal_gerado.png', conf=0.9):
        time.sleep(1)
    
    _salvar_pdf()
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Demonstrativo Mensal gerado']), nome=andamento)
    print('✔ Demonstrativo Mensal')
    
    p.press('esc', presses=4)
    time.sleep(2)


@_time_execution
def run():
    empresas = _open_lista_dados()
    andamentos = 'Faturamento X Compra'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
    
        if not _login(empresa, andamentos):
            continue
        relatorio_darf_dctf(empresa, andamentos)


if __name__ == '__main__':
    run()
