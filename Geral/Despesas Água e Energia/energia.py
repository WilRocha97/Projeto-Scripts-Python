# -*- coding: utf-8 -*-
import time, re, os
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def login(driver, email, senha):
    url = 'https://servicosonline.cpfl.com.br/agencia-webapp/#/login'
    driver.get(url)
    
    # aguarda o campo de email
    while not _find_by_id('documentoEmail', driver):
        time.sleep(1)
    
    data = {
        'url': url,
        'sitekey': '6LcJDCwUAAAAAPZsx3c7deGx7REdi5U3eNERQ_0j',
    }
    '<div class="recaptcha-checkbox-border" role="presentation" style="display: none;"></div>'
    token = _solve_recaptcha(data)
    
    # insere token
    script = 'document.getElementById("g-recaptcha-response").innerHTML="{}";'.format(str(token))
    driver.execute_script(script)
    
    # insere email
    driver.find_element(by=By.ID, value='documentoEmail').send_keys(email)
    
    # insere senha
    driver.find_element(by=By.ID, value='password').send_keys(senha)
    
    # clica para entrar
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/main/div/div[2]/div/div/div/div/div/div/div[3]/div[2]/div/div/div[1]/div/form/div/div[3]/div/button')\
        .click()
    

    
    time.sleep(22)
    
    
    return driver, resultado
        
        
@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    status, driver = _initialize_chrome(options)
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        nome, email, senha = empresa
        
        driver, resultado = login(driver, email, senha)
        if not resultado:
            continue
        
    
if __name__ == '__main__':
    run()
