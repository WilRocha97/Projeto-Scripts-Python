# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj):
    # entra no site
    driver.get('https://sindical.net/Sindical/Web/SeaacJundiai/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    while not _find_by_id('formLogin:i_cnpj', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='formLogin:i_cnpj').send_keys(cnpj)
    
    # clica em continuar
    driver.find_element(by=By.XPATH, value='/html/body/div[2]/form[1]/div[1]/div[2]/table/tbody/tr[3]/td/button/span[2]').click()
    time.sleep(1)
    
    try:
        avisos = re.compile(r'(CNPJ INVÁLIDO!)').search(driver.page_source).group(1)
        return driver, avisos
    except:
        return driver, ''


def gera_boleto(driver, cnpj, valor, data):
    while not _find_by_id('form_menu_principal_web:j_idt41', driver):
        time.sleep(1)
        
    print('>>> Gerando boletos')
    # clica em emitir boleto
    driver.find_element(by=By.XPATH, value='/html/body/table/tbody/tr[1]/td/form/table[1]/tbody/tr/td/div/ul/li[2]/a/span').click()
    
    while not _find_by_id('formWebContribuinte:j_idt65', driver):
        time.sleep(1)
        
    # clica em novo boleto
    driver.find_element(by=By.ID, value='formWebContribuinte:j_idt71').click()
    
    # espera a tela para inserir as infos do boleto aparecerem
    while not _find_by_id('formWebContribuinte:j_idt111', driver):
        time.sleep(1)
        
    # insere a data de referência
    driver.find_element(by=By.ID, value='formWebContribuinte:j_idt71').send_keys(data)
    
    # clica em adicionar
    driver.find_element(by=By.ID, value='formWebContribuinte:j_idt124').send_keys(data)
    
    
    
    
    
    return driver


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    data = empresa[5]
    
    print('100 - SEAAC - Sindicato Jundiaí e Região')
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj)
    
    # gera os boletos
    driver = gera_boleto(driver, cnpj, valor, data)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'
