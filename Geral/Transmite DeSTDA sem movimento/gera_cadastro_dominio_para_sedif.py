# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from dominio_comum import _login
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def arquivo_destda(empresa, periodo, andamento):
    cod, cnpj_limpo, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Informativo
    p.press('n')
    # Federais
    time.sleep(0.5)
    p.press('f')
    # DeSTDA
    time.sleep(0.5)
    if not _find_img('destda_opcao.png', conf=0.95):
        p.press('esc', presses=5)
        _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome,
                                         'Empresa não possuí opção de DeSTDA']),
                               nome=andamento)
        print('❗ Empresa não possuí opção de DeSTDA')
        return False
    p.press('d')
    time.sleep(0.5)
    p.press('d')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    
    while not _find_img('destda_janela.png', conf=0.9):
        time.sleep(1)
    
    p.write(periodo)
    
    if _find_img('arquivo_destda.png', conf=0.95):
        _click_img('arquivo_destda.png', conf=0.95)
    
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    p.write(f'M:\DeSTDA - {cnpj_limpo}.txt')
    
    time.sleep(0.5)
    
    p.hotkey('alt', 'p')
    
    while not _find_img('perfil_contribuinte.png', conf=0.9):
        time.sleep(1)
        if _find_img('cadastro_vazio.png', conf=0.95):
            p.press('enter')
    
    if _find_img('selecionar_dados.png', conf=0.95):
        _click_img('selecionar_dados.png', conf=0.95)
        time.sleep(0.5)
        p.write('0')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.write('0')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.write('1')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        
        p.hotkey('alt', 'g')
        time.sleep(1)
        
        '''if _find_img('PeriodoInvalido.png', conf=0.95):
            p.press('enter')
            time.sleep(0.5)
            p.press('esc')
            time.sleep(2)
            p.hotkey('alt', 'n')
            time.sleep(2)
            p.press('esc', presses=5)
            _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome, 'Erro na competência']),
                                   nome=andamento)
            print('❌ Erro na competência')
            return False'''
    
    p.press('esc')
    
    while not _find_img('destda_janela.png', conf=0.9):
        time.sleep(1)
    
    p.hotkey('alt', 'o')
    time.sleep(1)
    
    while not _find_img('final_da_exportacao.png', conf=0.9):
        time.sleep(1)
        
        """if _find_img('NaoExisteVigencia.png', conf=0.95):
            p.press('enter')
            time.sleep(0.5)
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome, 'Não existe vigência no período do arquivo']),
                                   nome=andamento)
            print('❌ Não existe vigência no período do arquivo')
            return False"""
        
        if _find_img('saldo_nao_calculado.png', conf=0.95):
            p.press('enter')
            time.sleep(0.5)
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome,
                                             'Saldo de impostos não foi apurado no período selecionado']),
                                   nome=andamento)
            print('❌ Saldo de impostos não foi apurado no período selecionado')
            return False
        
        """if _find_img('NaoExisteParametroVigencia.png', conf=0.95):
            p.press('enter')
            time.sleep(0.5)
            p.press('esc', presses=5)
            _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome,
                                             'Não existe parametro cadastrado para a competência']),
                                   nome=andamento)
            print('❌ Não existe parametro cadastrado para a competência')
            return False"""
    
    p.press('enter')
    time.sleep(1)
    p.press('esc')
    _escreve_relatorio_csv(';'.join([cod, cnpj_limpo, nome, 'Arquivo Gerado']), nome=andamento)
    print('✔ Arquivo Gerado')
    return True


@_time_execution
def run():
    periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos para importar cadastro de DeSTDA'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        
        if not _login(empresa, andamentos):
            continue
            
        arquivo_destda(empresa, periodo, andamentos)
            

if __name__ == '__main__':
    run()