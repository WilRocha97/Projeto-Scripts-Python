# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from sys import path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv
from chrome_comum import _initialize_chrome, _send_input


def _login(empresa, andamentos):
    cod, cnpj, nome = empresa
    # espera a tela inicial do domínio
    '''while p.locateOnScreen(r'imgs_c\inicial.png', confidence=0.5):
                    time.sleep(1)'''

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
        p.press('f8')
    
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
            while p.locateOnScreen(r'imgs_c\trocar_empresa.png', confidence=0.9):
                time.sleep(1)
            return False
        
        if p.locateOnScreen(r'imgs_c\fase_dois_do_cadastro.png', confidence=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if p.locateOnScreen(r'imgs_c\conforme_modulo.png', confidence=0.9):
            p.press('enter')
            time.sleep(1)

        if p.locateOnScreen(r'imgs_c\aviso_regime.png', confidence=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
    
    p.press('esc', presses=5)
    time.sleep(1)

    return True


def _login_web():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = _initialize_chrome()
    
    driver.get('https://www.dominioweb.com.br/')
    _send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', 'robo@veigaepostal.com.br', driver)
    _send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', 'Rb#0086*', driver)
    driver.find_element(by=By.ID, value='enterButton').click()
    
    caminho = os.path.join('imgs_c', 'abrir_app.png')
    caminho2 = os.path.join('imgs_c', 'abrir_app_2.png')
    while not p.locateOnScreen(caminho, confidence=0.9):
        if p.locateOnScreen(caminho2, confidence=0.9):
            sleep(1)
            while p.locateOnScreen(caminho2, confidence=0.9):
                p.click(p.locateCenterOnScreen(caminho2, confidence=0.9), button='left')
                return driver
        
        sleep(1)
    if p.locateOnScreen(caminho, confidence=0.9):
        p.click(p.locateCenterOnScreen(caminho, confidence=0.9), button='left')
        return driver


def _abrir_escrita_fiscal(driver):
    while not p.locateOnScreen(r'imgs_c/modulos.png', confidence=0.9):
        sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    sleep(3)
    driver.quit()
    p.click(p.locateCenterOnScreen(r'imgs_c/escrita_fiscal.png', confidence=0.9), button='left', clicks=2)
    while not p.locateOnScreen(r'imgs_c/login_modulo.png', confidence=0.9):
        sleep(1)
    
    p.moveTo(p.locateCenterOnScreen(r'imgs_c/insere_usuario.png', confidence=0.9))
    local_mouse = p.position()
    p.click(int(local_mouse[0] + 120), local_mouse[1], clicks=2)
    
    sleep(0.5)
    p.press('del', presses=10)
    p.write('robo')
    sleep(0.5)
    p.press('tab')
    sleep(0.5)
    p.press('del', presses=10)
    p.write('Rb#0086*')
    sleep(0.5)
    p.hotkey('alt', 'o')
    while not p.locateOnScreen(r'imgs_c/onvio.png', confidence=0.9):
        sleep(1)