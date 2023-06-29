# -*- coding: utf-8 -*-
import time, re, os
from selenium import webdriver
from selenium.webdriver.common.by import By

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice


@_time_execution
def run():
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    options.add_extension('C:/Users/willianr/AppData/Local/Google/Chrome/User Data/Default/Extensions/odncgaobfipblkjnmdpngmglgbngaabp/0.0.9_0.crx')
    
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        

if __name__ == '__main__':
    run()