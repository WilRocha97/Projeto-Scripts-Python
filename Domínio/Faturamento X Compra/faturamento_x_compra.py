# -*- coding: utf-8 -*-
import shutil, pyperclip, time, os, pyautogui as p
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
    # Confirmar
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('demonstrativo_mensal.png', conf=0.9):
        time.sleep(1)
    
    # configura o ano e digita no domínio
    ano = str(datetime.now().year)
    
    time.sleep(1)
    p.write(f'01{ano}')
    
    time.sleep(1)
    p.press('tab', presses=2)
    
    time.sleep(1)
    p.write(f'12{ano}')
    
    # gera o relatório
    time.sleep(1)
    p.hotkey('alt', 'o')
    
    # espera gerar
    while not _find_img('demonstrativo_mensal_gerado.png', conf=0.9):
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False
        time.sleep(1)

    guia = os.path.join('C:', 'Demonstrativo Mensal.pdf')
    while not os.path.exists(guia):
        _salvar_pdf()
    
    p.press('esc', presses=4)
    time.sleep(2)
    
    mover_demonstrativo(empresa, ano)
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Demonstrativo Mensal gerado']), nome=andamento)
    print('✔ Demonstrativo Mensal gerado')

def mover_demonstrativo(empresa, ano):
    os.makedirs('execução/Demonstrativos', exist_ok=True)
    cod, cnpj, nome = empresa
    execucoes = "V:\\Setor Robô\\Scripts Python\\Domínio\\Faturamento X Compra\\execução\\Demonstrativos"
    guia = os.path.join('C:', 'Demonstrativo Mensal.pdf')
    while not os.path.exists(guia):
        time.sleep(1)
    while os.path.exists(guia):
        try:
            arquivo = f'{cod} - {nome.replace("/", " ")} - Demonstrativo Mensal {ano}.pdf'
            shutil.move(guia, os.path.join(execucoes, arquivo))
            time.sleep(4)
        except:
            pass


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
