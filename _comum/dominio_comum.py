# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv


def _login(empresa, andamentos):
    cod, cnpj, nome = empresa
    # espera a tela inicial do domínio
    while p.locateOnScreen(r'imgs_c\inicial.png', confidence=0.9):
        time.sleep(1)

    # espera abrir a janela de seleção de empresa
    while not p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
        p.press('f8')
    while p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
        time.sleep(1)
    
    # clica para pesquisar empresa por código
    if p.locateOnScreen(r'imgs_c\codigo.png', confidence=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(cod)
    time.sleep(3)
    
    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
        time.sleep(1)
        if p.locateOnScreen(r'imgs_c\sem_parametro.png', confidence=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not p.locateOnScreen(r'imgs_c\parametros.png', confidence=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False
            
        if p.locateOnScreen(r'imgs_c\nao_existe_parametro.png', confidence=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro cadastrado para esta empresa']), nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
                time.sleep(1)
            return False
        
        if p.locateOnScreen(r'imgs_c\empresa_nao_usa_sistema.png', confidence=0.9) or p.locateOnScreen(r'imgs_c\empresa_nao_usa_sistema_2.png', confidence=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não está marcada para usar este sistema']), nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            while p.locateOnScreen(r'imgs_c\'trocar_empresa.png', confidence=0.9):
                time.sleep(1)
            return False
        
        if p.locateOnScreen(r'imgs_c\conforme_modulo.png', confidence=0.9):
            p.press('enter')
            time.sleep(1)
    
    p.press('esc', presses=5)

    return True
