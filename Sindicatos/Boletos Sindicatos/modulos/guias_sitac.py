# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _escreve_relatorio_csv
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj, senha):
    # entra no site
    driver.get('http://sitac.gladiador.net.br:8080/PortalSindical/login.jsf')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    while not _find_by_id('radio1', driver):
        time.sleep(1)
    time.sleep(1)
    
    # clica para abrir o menu
    driver.find_element(by=By.ID, value='radio1').click()
    time.sleep(1)
    
    # clica em CNPJ
    driver.find_element(by=By.ID, value='radio1_1').click()
    time.sleep(1)
    
    # insere CNPJ
    driver.find_element(by=By.ID, value='cnpjEmp').send_keys(cnpj)
    time.sleep(1)
    
    # insere Senha
    driver.find_element(by=By.ID, value='passCnpjEmp').send_keys(senha)
    time.sleep(1)
    
    # clica em continuar
    driver.find_element(by=By.ID, value='LogarCnpjEmp').click()
    time.sleep(1)
    
    while not _find_by_id('mn_dashboard', driver):
        time.sleep(1)
        
    # clica em continuar
    driver.find_element(by=By.ID, value='mn_financeiro').click()
    
    while not _find_by_id('tabelaConf', driver):
        time.sleep(1)
       
    return driver, ''

def gera_boleto(driver, cnpj, valor):
    print('>>> Verificando boletos')

    _escreve_relatorio_csv(f'{cnpj};{valor};Não tem nada aqui', nome='Boletos Sindicato')
    return driver


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    senha = empresa[7]
    
    print('8 - SITAC - Sindicato dos Trabalhadores nas Indústrias de Alimentação de Campinas')
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj, senha)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {avisos}'
    
    # gera os boletos
    driver = gera_boleto(driver, cnpj, valor)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'
