# -*- coding: utf-8 -*-
import time, re, os
from requests import Session
from bs4 import BeautifulSoup
from requests import Session
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def login():
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument('start-maximized')
    
    status, driver = _initialize_chrome(options)
    
    url = 'https://servicosonline.cpfl.com.br/agencia-webapp/#/via-pagamento'
    driver.get(url)
    
    data = {
        'url': url,
        'sitekey': '6LcJDCwUAAAAAPZsx3c7deGx7REdi5U3eNERQ_0j',
    }
    
    token = _solve_recaptcha(data)
    
    url = 'https://servicosonline.cpfl.com.br/agencia-webapi/api/captcha/token?encodedResponse=' + token
    driver.post(url)
    
    time.sleep(33)
    
    
def login_request():
    with Session() as s:
        # entra no site
        url = 'https://servicosonline.cpfl.com.br/agencia-webapp/#/via-pagamento'
        s.get(url)
        
        data = {
            'url': url,
            'sitekey': '6LcJDCwUAAAAAPZsx3c7deGx7REdi5U3eNERQ_0j',
        }
        
        token = _solve_recaptcha(data)
        
        s.post('https://servicosonline.cpfl.com.br/agencia-webapi/api/captcha/token?encodedResponse=' + token)
        
        res = s.get('https://servicosonline.cpfl.com.br/agencia-webapp/modules/via-pagamento/partials/modal-via-pagamento-login.html')
        
        soup = BeautifulSoup(res.content, 'html.parser')
        soup = soup.prettify()
        print(soup)
        time.sleep(33)
        
        
@_time_execution
def run():
    os.makedirs('execução/Boletos', exist_ok=True)
    # seleciona a lista de dados
    empresas = _open_lista_dados()
    
    # configura de qual linha da lista começar a rotina
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # para cada linha da lista executa
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, bla, bla, bla = empresa
        # printa o indice da lista
        _indice(count, total_empresas, empresa)
        login_request()
        
    
if __name__ == '__main__':
    run()
