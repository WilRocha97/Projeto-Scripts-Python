# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj, usuario, senha):
    # entra no site
    driver.get('http://metalcampinas.consir.com.br/index.php')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o botão para selecionar o login da empresa
    while not _find_by_path('/html/body/section/div/div[1]/div[2]', driver):
        time.sleep(1)
    time.sleep(1)
    
    # clica em empresa
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div[1]/div[2]').click()
    time.sleep(1)
    
    # espera campo para inserir o CNPJ
    while not _find_by_id('busca', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='busca').send_keys(cnpj)
    
    # clica em continuar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div[2]/form[1]/div[2]/input').click()
    time.sleep(1)
    
    # espera campo para inserir o CNPJ
    while not _find_by_path('/html/body/section/div/div/form/div[1]/input', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o usuario
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[1]/input').send_keys(usuario)
    
    # insere a senha
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[2]/input').send_keys(senha)
    
    # clica em continuar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[3]/input').click()
    
    return driver, ''


def gera_boleto(driver, valor):
    # espera aparecer o botão de gerar boleto
    while not _find_by_path('/html/body/header/div/nav/ul/li[2]/a', driver):
        time.sleep(1)
    time.sleep(1)
    
    driver.get('http://metalcampinas.consir.com.br/index.php?pg=novoboleto')
    
    # espera aparecer o dropdown
    while not _find_by_path('/html/body/section/div/div/form[2]/div[2]/div[2]/span/span[1]/span/span[1]', driver):
        time.sleep(1)
    
    # clica no dropdown
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form[2]/div[2]/div[2]/span/span[1]/span/span[1]').click()
    time.sleep(2)
    
    # clica no iten do dropdown
    driver.find_element(by=By.XPATH, value='/html/body/span/span/span[2]/ul/li[2]').click()
    time.sleep(1)
    
    # espera aparecer a tela para inserir o valor
    while not _find_by_id('valor', driver):
        time.sleep(1)
    
    # insere o valor
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form[2]/div[5]/div[2]/input').send_keys(valor)
    
    time.sleep(1)
    
    # clica em gerar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form[2]/div[6]/div/input').click()
    time.sleep(1)
    
    return driver, ''


def salva_boleto(driver, cnpj):
    cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
    # espera aparecer o botão de imprimir
    while not _find_by_path('/html/body/section/div/div[2]/a[1]', driver):
        time.sleep(1)
    
    # clica em imprimir
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div[2]/a[1]').click()
    time.sleep(1)
    
    while not _find_img('imprimir.png', conf=0.9):
        time.sleep(1)
    
    _click_img('imprimir.png', conf=0.9, timeout=1)
    
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
    
    # nome do arquivo
    p.write(f'Boleto 11-SINTICOM - {cnpj_limpo}.pdf')
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
    
    return driver, ''


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    usuario = empresa[6]
    senha = empresa[7]
    
    print('10 - METALCAMPINAS - Sindicato dos Trabalhadores nas Indústrias Metalúrgicas, Mecatrônicas, Materiais Elétricos e Fibras Óticas de Campinas e Região')
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj, usuario, senha)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {avisos}'
    
    # gera os boletos
    driver, avisos_2 = gera_boleto(driver, valor)
    driver, avisos_3 = salva_boleto(driver, cnpj)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'
