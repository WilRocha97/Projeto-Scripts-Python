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


def login(driver, cnpj):
    # entra no site
    driver.get('https://boletos.sindpd.org.br/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    
    return driver, avisos

    
def run(cnpj, valor, usuario, senha, funcionarios):
    print('28 - SINDPD - Sindicato dos Trabalhadores em Processamento de Dados e Tecnologia da Informação do Estado de São Paulo')
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    driver, avisos = login(driver, cnpj)
    if not avisos:
        driver = gera_boleto(driver, cnpj, valor)
    else:
        _escreve_relatorio_csv(f'{cnpj};{valor}', nome='Boletos Sindicato')
        print(f"❌ {avisos}")
            
    driver.close()
