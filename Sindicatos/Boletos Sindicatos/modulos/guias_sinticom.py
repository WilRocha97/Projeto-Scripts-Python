# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj, usuario, senha):
    # entra no site
    driver.get('http://sinticomcampinas.consir.com.br/index.php')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o botão para selecionar o login da empresa
    while not _find_by_path('/html/body/section/div/div[1]/div[2]', driver):
        time.sleep(1)
    time.sleep(1)
    
    # clica em empresa
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div[1]/div[2]').click()
    time.sleep(1)
    
    # espera campo para inserir o CNPJ
    while not _find_by_id('busca', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='busca').send_keys(cnpj)
    
    # clica em continuar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div[2]/form[1]/div[2]/input').click()
    time.sleep(1)
    
    # espera campo para inserir o CNPJ
    while not _find_by_path('/html/body/section/div/div/form/div[1]/input', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o usuario
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[1]/input').send_keys(usuario)
    
    # insere a senha
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[2]/input').send_keys(senha)
    
    # clica em continuar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form/div[3]/input').click()
    
    return driver, ''

    
def gera_boleto(driver, valor):
    # espera aparecer o botão de gerar boleto
    while not _find_by_path('/html/body/header/div/nav/ul/li[2]/a', driver):
        time.sleep(1)
    time.sleep(1)
    
    driver.get('http://sinticomcampinas.consir.com.br/index.php?pg=novoboleto')
    
    # espera aparecer o dropdown
    while not _find_by_path('/html/body/section/div/div/form[2]/div[2]/div[2]/span/span[1]/span/span[1]', driver):
        time.sleep(1)

    # clica no dropdown
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form[2]/div[2]/div[2]/span/span[1]/span/span[1]').click()
    time.sleep(2)

    # clica no iten do dropdown
    driver.find_element(by=By.XPATH, value='/html/body/span/span/span[2]/ul/li[2]').click()
    time.sleep(1)
    
    # espera aparecer a tela para inserir o valor
    while not _find_by_id('valor', driver):
        time.sleep(1)
        
    # insere o valor
    driver.find_element(by=By.id, value='/html/body/span/span/span[2]/ul/li[2]').send_keys(valor)
    time.sleep(1)
    
    # clica em gerar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/form[2]/div[6]/div/input').click()
    time.sleep(1)
    
    return driver, ''


def salva_boleto(cnpj):
    # nome do arquivo
    p.write(f'Boleto 11-SINTICOM - {cnpj}.pdf')
    return True
    
    
def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    usuario = empresa[6]
    senha = empresa[7]
    
    print('11 - SINTICOM - Sindicato dos Trabalhadores nas Indústrias da Construção e do Mobiliário de Campinas e Região')
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj, usuario, senha)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {avisos}'
    
    # gera os boletos
    driver = gera_boleto(driver, valor)
    resultado = salva_boleto(cnpj)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'
