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
    driver.get('https://www.mirah.com.br/localizacnpj.aspx?CodigoEntidade=000020399897920&Contribuicao=Assistencial')
    print('>>> Acessando o site')
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
    
    print('69 - SINDITERCEIRIZADOS - Sindicato dos Empregados e Trabalhadores nas Empresas de Prestação de Serviços de Asseio e Conservação Limpeza Urbana, Limpeza Ambiental, Áreas Verdes dos Municípios de Jundiaí e Região.')
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
