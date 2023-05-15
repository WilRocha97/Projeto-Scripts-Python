# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj):
    # entra no site
    driver.get('https://www.mirah.com.br/localizacnpj.aspx?CodigoEntidade=020818905292&Contribuicao=QUOTA%20DE%20PARTICIPACAO%20NEGOCIAL')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    while not _find_by_id('CnpjTextBox', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='CnpjTextBox').send_keys(cnpj)
    
    # clica em continuar
    driver.find_element(by=By.ID, value='cmdBuscar').click()
    
    try:
        avisos = re.compile(r'(CNPJ não Cadastrado. Por favor, insira os dados.)').search(driver.page_source).group(1)
        return driver, avisos
    except:
        return driver, False


def gera_boleto(driver, cnpj, valor):
    print('>>> Gerando boletos')
    # insere o valor do boleto
    driver.find_element(by=By.ID, value='txtValor').send_keys(valor)
    
    # clica em emitir
    driver.find_element(by=By.ID, value='cmdConfirmar').click()
    
    os.makedirs('execução/Boletos', exist_ok=True)
    
    while not _find_img('imprimir.png', conf=0.9):
        time.sleep(1)
    
    _click_img('imprimir.png', conf=0.9)
    
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
    
    # nome do arquivo
    p.write(f'Boleto 39-SINTHOJUR - {cnpj}.pdf')
    time.sleep(0.5)
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Sindicatos\Boletos Sindicatos\execução\Boletos')
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(1)
    
    return driver
  
    
def run(cnpj, valor, usuario, senha, funcionarios):
    print('39 - SINTHOJUR - Sindicato dos Trabalhadores em Hotéis, Motéis, Restaurantes, Bares, Lanchonetes e Fast-food de Jundiaí e Região')
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    driver, avisos = login(driver, cnpj)
    if not avisos:
        driver = gera_boleto(driver, cnpj, valor)
        driver.close()
        return '✔ Boleto gerado'
    else:
        driver.close()
        return f'❌ {avisos}'
