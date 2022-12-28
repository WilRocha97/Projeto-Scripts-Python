# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from time import sleep
from datetime import datetime
from sys import path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv
from chrome_comum import _initialize_chrome, _send_input_xpath, _find_by_class
from pyautogui_comum import _find_img, _wait_img, _click_img


def _login(empresa, andamentos):
    cod, cnpj, nome = empresa
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')
    
    sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(cod)
    sleep(3)
    
    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                sleep(1)
            p.press('esc')
            sleep(1)
            return False
            
        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro cadastrado para esta empresa']), nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                sleep(1)
            return False
        
        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img('empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não está marcada para usar este sistema']), nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            sleep(1)
            p.press('esc')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                sleep(1)
            return False
        
        if _find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            sleep(1)
            p.press('esc', presses=5, interval=1)
            _login(empresa, andamentos)
    
    if not verifica_empresa(cod):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não encontrada']), nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False
    
    p.press('esc', presses=5)
    sleep(1)

    return True


def verifica_empresa(cod):
    p.click(1280,82)

    time.sleep(1)
    p.hotkey('ctrl', 'c')
    p.hotkey('ctrl', 'c')
    cnpj_codigo = pyperclip.paste()
    
    time.sleep(0.5)
    codigo = cnpj_codigo.split('-')
    codigo = str(codigo[1])
    codigo = codigo.replace(' ', '')
    
    if codigo != cod:
        return False
    
    else:
        return True
    

def _login_web():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = _initialize_chrome()
    
    driver.get('https://www.dominioweb.com.br/')
    _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', 'robo@veigaepostal.com.br', driver)
    _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', 'Rb#0086*', driver)
    driver.find_element(by=By.ID, value='enterButton').click()
    
    caminho = os.path.join('imgs_c', 'abrir_app.png')
    caminho2 = os.path.join('imgs_c', 'abrir_app_2.png')
    while not p.locateOnScreen(caminho, confidence=0.9):
        if _find_by_class('trta1-btn-primary', driver):
            driver.find_element(by=By.CLASS_NAME, value='trta1-btn-primary').click()
        if p.locateOnScreen(caminho2, confidence=0.9):
            sleep(1)
            while p.locateOnScreen(caminho2, confidence=0.9):
                p.click(p.locateCenterOnScreen(caminho2, confidence=0.9), button='left')
        
    if p.locateOnScreen(caminho, confidence=0.9):
        p.click(p.locateCenterOnScreen(caminho, confidence=0.9), button='left')

    while not locateOnScreen(path.join('imgs_c', 'modulos.png'), confidence=0.9):
        sleep(1)
    driver.quit()
    return True


def _abrir_modulo(modulo):
    while not p.locateOnScreen(r'imgs_c/modulos.png', confidence=0.9):
        sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    sleep(1)
    p.click(p.locateCenterOnScreen(r'imgs_c/modulo' + modulo + '.png', confidence=0.9), button='left', clicks=2)
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


def _salvar_pdf():
    p.hotkey('ctrl', 'd')
    while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
    
    if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9) or _find_img('cliente_m.png', pasta='imgs_c', conf=0.9):
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)
                
        _click_img('cliente_m.png', pasta='imgs_c', conf=0.9)
        _click_img('cliente_c.png', pasta='imgs_c', conf=0.9)
        time.sleep(5)
    
    p.press('enter')
    
    timer = 0
    while not _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        if _find_img('erro_pdf.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f4')
            
        if _find_img('substituir.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('adobe.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
        time.sleep(1)
        timer += 1
        if timer > 30:
            _click_img('salvar_pdf.png', pasta='imgs_c', conf=0.9)
            while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            
            if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
                while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9):
                    _click_img('botao.png', pasta='imgs_c', conf=0.9)
                    time.sleep(3)
                _click_img('cliente_c.png', pasta='imgs_c', conf=0.9)
                time.sleep(5)
            
            p.press('enter')
            timer = 0

    while _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        p.hotkey('alt', 'f4')
        time.sleep(3)
    
    while _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
        p.press('esc')
        
    return True
