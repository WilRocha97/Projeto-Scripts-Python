# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path, _send_input
from comum_comum import _escreve_relatorio_csv
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj):
    print('>>> Acessando o site')
    # entra no site
    timer = 0
    while not _find_by_id('TextBoxCNPJ', driver):
        # abre o site da consulta e caso de erro é porque o site demorou pra responder,
        # nesse caso retorna um erro para tentar novamente
        try:
            driver.get('http://sisnet.sindesporte.com.br/sisnet40/negocial.aspx')
        except:
            return driver, False
        time.sleep(1)
        timer += 0
        if timer > 120:
            return driver, False
    
    _send_input('TextBoxCNPJ', cnpj, driver)
    time.sleep(1)
    
    # aperta ok para entrar
    driver.find_element(by=By.ID, value='ButtonOk').click()

    return driver, True

def gera_boleto(driver, valor, funcionarios):
    print('>>> Verificando boletos')
    # aguarda o botão de nova guia aparecer
    while not _find_by_id('ButtonNovoBoleto', driver):
        time.sleep(1)
    
    # aperta botão de nova guia
    driver.find_element(by=By.ID, value='ButtonNovoBoleto').click()
    
    print('>>> Aguardando tela para gerar o boleto')
    # aguarda o campo de quantidade de funcionários
    while not _find_by_id('GridViewGuias_TextBoxEmpregados_0', driver):
        time.sleep(1)
    
    print('>>> Gerando o boleto')
    # insere o número de funcionarios
    _send_input('GridViewGuias_TextBoxEmpregados_0', funcionarios, driver)
    time.sleep(1)
    # insere o valor do boleto
    _send_input('GridViewGuias_TextBoxValor_0', valor, driver)
    time.sleep(1)
    
    # aperta botão de gerar o boleto
    driver.find_element(by=By.ID, value='ButtonGerarGuia1').click()
    time.sleep(33)
    
    print('>>> Boleto gerado')
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
    if not resultados:
        driver.close()
        return f'❌ Erro no login - {resultados}'
    
    # gera os boletos
    driver, resultados = gera_boleto(driver, valor, funcionarios)
    driver.close()
    return f'✔ Boleto gerado - {resultados}'
