# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _escreve_relatorio_csv


def _login(empresa, andamentos):
    cod, cnpj, nome = empresa
    # espera a tela inicial do domínio
    _wait_img('Inicial.png', conf=0.9, timeout=-1)

    # espera abrir a janela de seleção de empresa
    while not _find_img('TrocarEmpresa.png', conf=0.9):
        p.press('f8')
    _wait_img('TrocarEmpresa.png', conf=0.9, timeout=-1)
    
    # clica para pesquisar empresa por código
    if _find_img('Codigo.png', conf=0.9):
        _click_img('Codigo.png', conf=0.9)
    p.write(cod)
    time.sleep(3)
    
    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('TrocarEmpresa.png', conf=0.9):
        time.sleep(1)
        if _find_img('NaoExisteParametro.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro cadastrado para esta empresa']), nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('TrocarEmpresa.png', conf=0.9):
                time.sleep(1)
            return False
        
        if _find_img('EmpresaNaoUsaSistema.png', conf=0.9) or _find_img('EmpresaNaoUsaSistema2.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não está marcada para usar este sistema']), nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            while _find_img('TrocarEmpresa.png', conf=0.9):
                time.sleep(1)
            return False
        
        if _find_img('ConformeModulo.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            
    return True
