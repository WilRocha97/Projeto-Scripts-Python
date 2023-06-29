# -*- coding: utf-8 -*-
import time, re, os
from pyautogui import press

import PIL.PdfParser
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Confere IR.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


def login(driver):
    driver.get('https://portal.conferironline.com.br/login?returnUrl=%2Fdownloads-center')
    email = '/html/body/app-root/app-full-layout/div/app-login-page/app-login/div/div/div/div[2]/div/app-form/div/div/form/app-overlay-loading/div/div/app-input[1]/div/input'
    senha = '/html/body/app-root/app-full-layout/div/app-login-page/app-login/div/div/div/div[2]/div/app-form/div/div/form/app-overlay-loading/div/div/app-input[2]/div/input'
    while not _find_by_path(email, driver):
        time.sleep(1)
    
    driver.find_element(by=By.XPATH, value=email).send_keys(user[0])
    driver.find_element(by=By.XPATH, value=senha).send_keys(user[1])
    time.sleep(1)
    
    driver.find_element(by=By.XPATH,
                        value='/html/body/app-root/app-full-layout/div/app-login-page/app-login/div/div/div/div[2]/div/app-form/div/div/form/app-overlay-loading/div/button')\
                        .click()
    time.sleep(2)
    return driver


def consulta(driver, cpf):
    driver.get('https://portal.conferironline.com.br/customers')
    pesquisa = '/html/body/app-root/app-content-layout/div/div/app-header/div/div/div[2]/ul/li[1]/app-search-bar/form/div/input'
    while not _find_by_path(pesquisa, driver):
        time.sleep(1)
    
    driver.find_element(by=By.XPATH, value=pesquisa).send_keys(cpf)
    time.sleep(2)
    press('enter')
    time.sleep(5)
    if not re.compile(r'class=\"f-w-600\">Cpf/Cnpj:</span> ' + cpf +' ').search(driver.page_source):
        return driver, 'Cliente não encontrado'
    else:
        url_cliente = re.compile(r'iconNewTab\" href=\"/(.+)\"><app-icon _ngcontent-\w\w\w-c115').search(driver.page_source).group(1)
        cliente_id = url_cliente.split('=')[1]
        
    driver.get('https://portal.conferironline.com.br/customer-profile/actions?customerId=' + cliente_id)
    time.sleep(3)
    
    driver.find_element(by=By.XPATH,
                        value='/html/body/app-root/app-content-layout/div/div/div/div[2]/main/div/div/div/app-customer-profile-page/app-profile-customer/app-overlay-loading/div/app-profile-tabset/app-overlay-loading/div/div/div[2]/app-actions-ecac/app-card/div/div/div/div[3]/a')\
                        .click()
    
    abas = driver.window_handles
    driver.switch_to.window(abas[1])
    time.sleep(22)
    return driver, ''
    

@_time_execution
def run():
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    options.add_extension('V:/Setor Robô/Scripts Python/Ecac/Confe IR/ignore/0.0.9_0.crx')
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver = login(driver)
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        cpf, nome = empresa
        
        driver, resultado = consulta(driver, cpf)
        print(resultado)
        

if __name__ == '__main__':
    run()