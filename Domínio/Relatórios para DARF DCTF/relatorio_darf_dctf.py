# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login


def relatorio_darf_dctf(empresa, periodo, andamento):
    cod, cnpj, nome = empresa
    _wait_img('Relatorios.png', conf=0.9, timeout=-1)
    # Relatóriosm
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('i')
    # Resumo
    time.sleep(0.5)
    p.press('m')

    while not _find_img('ResumoDeImpostos.png', conf=0.9):
        time.sleep(1)
        if _find_img('SemParametroVigencia.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro']), nome=andamento)
            print('❌ Não existe parametro')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

    # Período
    p.write(periodo)
    p.press('tab')
    time.sleep(1)
    p.write(periodo)
    time.sleep(0.5)

    # Todos os impostos
    p.hotkey('alt', 't')
    time.sleep(1)

    while _find_img('DestacarLinhas.png', conf=0.95):
        _click_img('DestacarLinhas.png', conf=0.95)
        time.sleep(0.5)
        
    '''while _find_img('DetalharDados.png', conf=0.95):
                    _click_img('DetalharDados.png', conf=0.95)
                    time.sleep(0.5)'''

    p.hotkey('alt', 'o')
    time.sleep(1)
    sem_layout = 0

    while not _find_img('ResumoGerado.png', conf=0.9):

        if _find_img('ImpostoSemLayout.png', conf=0.9):
            sem_layout = 1
            p.press('enter')
        if sem_layout == 1:
            while not _find_img('ResumoDeImpostos.png', conf=0.9):
                p.press('enter')
                time.sleep(1)
            p.press('esc', presses=4)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Imposto sem layout']), nome=andamento)
            print('❌ Imposto sem layout')
            return False

        time.sleep(3)
        if _find_img('SemDados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('SemImposto.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('ResumoCalculado.png', conf=0.9):
            break

    _click_img('SalvarPDF.png', conf=0.9)
    while not _find_img('SalvarEmPDF.png', conf=0.9):
        time.sleep(1)

    if not _find_img('ClienteCSelecionado.png', conf=0.9):
        _click_img('Botao.png', conf=0.9)
        _wait_img('ClienteC.png', conf=0.9, timeout=-1)
        _click_img('ClienteC.png', conf=0.9)
        time.sleep(5)

    p.press('enter')

    while not _find_img('PDFAberto.png', conf=0.9):
        if _find_img('Substituir.png', conf=0.9):
            p.hotkey('alt', 'y')
        time.sleep(1)

    p.hotkey('alt', 'f4')
    time.sleep(2)

    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório gerado']), nome=andamento)
    print('✔ Relatório gerado')

    p.press('esc', presses=3)


@_time_execution
def run():
    periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/0000')
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
        relatorio_darf_dctf(empresa, periodo, andamentos)


if __name__ == '__main__':
    run()
