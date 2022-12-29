# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from PIL import Image
from sys import path
import os, time

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

os.makedirs('execução/documentos', exist_ok=True)


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def login(options, cnpj, nome):
    status, driver = _initialize_chrome(options)
    
    print('>>> Consultando cnpj', cnpj)
    url = 'https://conectividadesocialv2.caixa.gov.br/sicns/'
    
    driver.get(url)
    time.sleep(546)
    
@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        
        _indice(count, total_empresas, empresa)
        
        if not login(options, cnpj, nome):
            break


if __name__ == '__main__':
    run()
