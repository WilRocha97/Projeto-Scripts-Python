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
    driver.get('http://sisnet.sindesporte.com.br/sisnet40/negocial.aspx')
    print('>>> Acessando o site')
    time.sleep(1)

    return driver, ''

def gera_boleto(driver, cnpj, valor, funcionarios):
    print('>>> Verificando boletos')

    
    return driver, resultados


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    funcionarios = empresa[8]
    
    print('17 - SINDESPORTE - Sindicato dos Empregados de Clubes Esportivos e em Federações, Confederações e Academias Esportivas no Estado de São Paulo')
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, resultados = login(driver, cnpj)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {resultados}'
    
    # gera os boletos
    driver, resultados = gera_boleto(driver, cnpj, valor, funcionarios)
    driver.close()
    return f'✔ Boleto gerado - {resultados}'
