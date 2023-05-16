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
    driver.get('https://boletos.sindpd.org.br/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    while not _find_by_id('CnpjEmpresa', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='CnpjEmpresa').send_keys(cnpj)
    
    # insere a senha da empresa
    driver.find_element(by=By.ID, value='SenhaEmpresa').send_keys(senha)
    
    # clica em continuar
    driver.find_element(by=By.ID, value='btacessa').click()
    
    while not _find_by_path('/html/body/div[4]/h2', driver):
        try:
            # mude para o contexto do alerta
            alert = driver.switch_to.alert
            # confirme o alerta
            alert.accept()
        except:
            pass
        
    print('>>> Verificando boletos')
    
    avisos = re.compile('ATENÇÃO! (.+).<br> ').search(driver.page_source)
    if avisos:
        avisos = avisos.group(1)
        avisos = avisos.replace('<br>', '')
        return driver, avisos
    
    return driver, ''

def gera_boleto(driver, cnpj, valor, data, funcionarios, responsavel, email):
    data = data.split('-')
    data = f'{data[1]}/{data[0]}'
    # clica em contribuição assistencial
    driver.find_element(by=By.ID, value='contrib2').click()
    
    while not _find_by_path('/html/body/div[4]/b[2]/font/div[2]/table', driver):
        time.sleep(1)
    
    print(driver.page_source)
    _escreve_relatorio_csv(f'{cnpj};{valor};Não tem nada aqui', nome='Boletos Sindicato')
    return driver


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    data = empresa[5]
    senha = empresa[7]
    funcionarios = empresa[8]
    responsavel = empresa[9]
    email = empresa[10]
    
    print('28 - SINDPD - Sindicato dos Trabalhadores em Processamento de Dados e Tecnologia da Informação do Estado de São Paulo')
    
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")

    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj, senha)
    
    # gera os boletos
    driver = gera_boleto(driver, cnpj, valor, data, funcionarios, responsavel, email)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'

