# -*- coding: utf-8 -*-
import pyperclip, time, os, subprocess, pyautogui as p
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

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Domínio.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


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
            sleep(1)
            p.press('esc')
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
            p.press('esc', presses=5)
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

        if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
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
    erro = 'sim'
    while erro == 'sim':
        try:
            p.click(1258,82)
    
            while True:
                try:
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    cnpj_codigo = pyperclip.paste()
                    break
                except:
                    pass
        
            time.sleep(0.5)
            codigo = cnpj_codigo.split('-')
            codigo = str(codigo[1])
            codigo = codigo.replace(' ', '')
        
            if codigo != cod:
                print(codigo)
                print(cod)
                return False
            else:
                return True
        except:
            erro = 'sim'
    

def _login_web():
    if not _find_img('app_controler.png', pasta='imgs_c', conf=0.9):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        status, driver = _initialize_chrome(options)
    
        driver.get('https://www.dominioweb.com.br/')
        _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', user[0], driver)
        _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', user[3], driver)
        driver.find_element(by=By.ID, value='enterButton').click()
        
        caminho = os.path.join('abrir_app.png')
        caminho2 = os.path.join('abrir_app_2.png')
        while not _find_img(caminho, pasta='imgs_c', conf=0.9):
            if _find_by_class('trta1-btn-primary', driver):
                driver.find_element(by=By.CLASS_NAME, value='trta1-btn-primary').click()
            if _find_img(caminho2, pasta='imgs_c', conf=0.9):
                sleep(1)
                while _find_img(caminho2, pasta='imgs_c', conf=0.9):
                    _click_img(caminho2, pasta='imgs_c', conf=0.9)
                break
                
        if _find_img(caminho, pasta='imgs_c', conf=0.9):
            _click_img(caminho, pasta='imgs_c', conf=0.9)
        
        print('>>> Aguardando modulos')
        while not _find_img('modulos.png', pasta='imgs_c', conf=0.9):
            sleep(1)
        driver.quit()
        return True
    else:
        _click_img('app_controler_desfocado.png', pasta='imgs_c', conf=0.9)
        sleep(2)
        if _find_img('lista_de_programas.png', pasta='imgs_c', conf=0.9):
            p.press('right', presses=2, interval=1)
            sleep(1)
            p.press('enter')
            

def _abrir_modulo(modulo):
    if _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        return True
    
    modulo_nome = ''
    if modulo == 'escrita_fiscal':
        modulo_nome = 'Escrita Fiscal'
    elif modulo == 'folha':
        modulo_nome = 'Folha'
    elif modulo == 'conexoes':
        modulo_nome = 'Conexões'
        
    print(f'>>> Abrindo modulo {modulo_nome}\n')
    while not _find_img('modulos.png', pasta='imgs_c', conf=0.9):
        sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    sleep(1)
    _click_img('modulo_' + modulo + '.png', pasta='imgs_c', conf=0.9, button='left', clicks=2)
    while not _find_img('login_modulo.png', pasta='imgs_c', conf=0.9):
        sleep(1)
    
    p.moveTo(_find_img('insere_usuario.png', pasta='imgs_c', conf=0.9))
    local_mouse = p.position()
    p.click(int(local_mouse[0] + 120), local_mouse[1], clicks=2)
    
    sleep(0.5)
    p.press('del', presses=10)
    p.write(user[1])
    sleep(0.5)
    p.press('tab')
    sleep(0.5)
    p.press('del', presses=10)
    p.write(user[2])
    sleep(0.5)
    p.hotkey('alt', 'o')
    while not _find_img('onvio.png', pasta='imgs_c', conf=0.9):
        sleep(1)

    time.sleep(5)
    
    if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
        p.hotkey('alt', 'o')
    return True


def _salvar_pdf():
    p.hotkey('ctrl', 'd')
    timer = 0
    while not _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False
    
    if not _find_img('cliente_c_selecionado.png', pasta='imgs_c', conf=0.9):
        while not _find_img('cliente_c.png', pasta='imgs_c', conf=0.9) or _find_img('cliente_m.png', pasta='imgs_c', conf=0.9):
            _click_img('botao.png', pasta='imgs_c', conf=0.9)
            time.sleep(3)
                
        _click_img('cliente_m.png', pasta='imgs_c', conf=0.9, timeout=1)
        _click_img('cliente_c.png', pasta='imgs_c', conf=0.9, timeout=1)
        time.sleep(5)
    
    p.press('enter')
    
    timer = 0
    while not _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        if _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
            p.press('esc')
            time.sleep(2)
            return False
        
        if _find_img('erro_pdf.png', pasta='imgs_c', conf=0.9) or _find_img('erro_pdf_2.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f4')
            
        if _find_img('substituir.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('adobe.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
        time.sleep(1)
        timer += 1
        if timer > 30:
            p.hotkey('ctrl', 'd')
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
        time.sleep(2)
        
    return True


def _verifica_dominio():
    
    if not _find_img('app_controler.png', pasta='imgs_c', conf=0.9):
        return 'dominio fechou'
    if _find_img('modulos.png', pasta='imgs_c', conf=0.9):
        return 'modulo fechou'
    else:
        return 'continue'


def _encerra_dominio():
    processo = 'C:\Program Files (x86)\GraphOn\AppController\AppController.exe'
    with open(os.devnull, 'w') as devnull:
        try:
            subprocess.call(['kill', str(processo)], stdout=devnull, stderr=subprocess.STDOUT)
            print('\nDomínio Web finalizado.\n')
        except:
            print('\nDomínio Web não finalizado.\n')
