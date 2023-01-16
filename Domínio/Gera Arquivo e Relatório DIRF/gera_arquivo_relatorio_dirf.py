# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login, _salvar_pdf


def dirf(empresa, ano, andamento):
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('n')
    # federais
    time.sleep(0.5)
    p.press('f')
    # DIRF
    time.sleep(0.5)
    p.press('i')
    # apartir de 2010
    time.sleep(0.5)
    p.press('p')
    
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

    if not relatorio_dirf(empresa, ano, andamento):
        return False
    
    
def relatorio_dirf(empresa, ano, andamento):
    cod, cnpj, nome = empresa
    
    # seleciona para gerar o relatório
    if _find_img('formulario.png'):
        _click_img('formulario.png')
    time.sleep(1)
    
    # seleciona para gerar com folha de pagamento
    if _find_img('folha_de_pagamento.png'):
        _click_img('folha_de_pagamento.png')
    time.sleep(1)
    
    # aabre a janela de outros dados
    p.hotkey('alt', 'o')
    
    while not _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)

    if _find_img('folha_de_pagamento.png', conf=0.95):
        p.press('esc')
        time.sleep(1)
        p.hotkey('alt', 'n')
        time.sleep(1)
        p.press('esc')
        
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não é possível editar aba de folha de pagamento']), nome=andamento)
        print('❌ Não é possível editar aba de folha de pagamento')
        return False
    
    _click_img('tem_folha.png', conf=0.95)
    time.sleep(2)
    
    # clicar para gerar de todos os colaboradores
    if _find_img('todos.png', conf= 0.95):
        _click_img('todos.png', conf=0.95)
    time.sleep(1)
    
    if _find_img('gerar_info_complementar.png', conf= 0.95):
        _click_img('gerar_info_complementar.png', conf=0.95)
    time.sleep(1)
    
    if _find_img('limitar_60_caractere.png', conf=0.95):
        _click_img('limitar_60_caractere.png', conf=0.95)
    time.sleep(1)
    
    p.hotkey('alt', 'g')
    
    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'o')
    
    
    # under construction
    
    
    _salvar_pdf()
    
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
        dirf(empresa, ano, andamentos)


if __name__ == '__main__':
    run()
