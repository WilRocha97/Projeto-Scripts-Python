# -*- coding: utf-8 -*-
import re, shutil, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf


def relatorio_darf_dctf(empresa, periodo, andamento):
    cod, cnpj, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('i')
    # Resumo
    time.sleep(0.5)
    p.press('m')
    verificacao = 'continue'
    while not _find_img('resumo_de_impostos.png', conf=0.9):
        
        if verificacao != 'continue':
            return verificacao
        time.sleep(1)
        if _find_img('vigencia_sem_parametro.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro']), nome=andamento)
            print('❌ Não existe parametro')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok'

    # Período
    p.write(periodo)
    p.press('tab')
    time.sleep(1)
    
    if _find_img('sem_parametro_vigencia.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro para a vigência: ' + periodo]), nome=andamento)
        print('❌ Não existe parametro para a vigência: ' + periodo)
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        p.press('esc')
        time.sleep(1)
        return 'ok'
    
    p.write(periodo)
    time.sleep(0.5)

    # Todos os impostos
    p.hotkey('alt', 't')
    time.sleep(1)

    while _find_img('destacar_linhas.png', conf=0.95):
        _click_img('destacar_linhas.png', conf=0.95)
        time.sleep(0.5)
        
    '''while _find_img('detalhar_dados.png', conf=0.95):
                    _click_img('detalhar_dados.png', conf=0.95)
                    time.sleep(0.5)'''

    p.hotkey('alt', 'o')
    time.sleep(1)
    sem_layout = 0

    while not _find_img('resumo_gerado_3.png', conf=0.8):
        if verificacao != 'continue':
            return verificacao
        time.sleep(1)
        if _find_img('imposto_sem_layout.png', conf=0.9):
            sem_layout = 1
            p.press('enter')
        if sem_layout == 1:
            while not _find_img('resumo_de_impostos.png', conf=0.9):
                if verificacao != 'continue':
                    return verificacao
                p.press('enter')
                time.sleep(1)
            p.press('esc', presses=4)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Imposto sem layout']), nome=andamento)
            print('❌ Imposto sem layout')
            return 'ok'

        time.sleep(3)
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok'

        if _find_img('sem_imposto.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok'

        if _find_img('resumo_calculado.png', conf=0.9) or _find_img('resumo_gerado_2.png', conf=0.9):
            break

    _salvar_pdf()
    mover_relatorio(cod)
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório gerado']), nome=andamento)
    print('✔ Relatório gerado')

    p.press('esc', presses=4)
    time.sleep(2)
    
    return 'ok'


def mover_relatorio(cod):
    folder = 'C:\\'
    for arq in os.listdir(folder):
        if arq.endswith('.pdf'):
            if re.compile(r'(.+) - .+\.pdf').search(arq).group(1) == f'Empresa {cod}':
                os.makedirs('execução/Relatórios', exist_ok=True)
                final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Relatórios para DARF DCTF\\execução\\Relatórios'
                shutil.move(os.path.join(folder, arq), os.path.join(final_folder, arq))


@_time_execution
@_barra_de_status
def run(window):
    periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Relatórios para DARF DCTF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]

    _login_web()
    _abrir_modulo('escrita_fiscal')

    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count + index)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index))} Restantes')
        _indice(count, total_empresas, empresa, index)

        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado = relatorio_darf_dctf(empresa, periodo, andamentos)
            
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
            
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break


if __name__ == '__main__':
    run()
