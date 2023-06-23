# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, usuario, senha):
    # entra no site
    driver.get('http://sinecoll.ddns.net:7070/Sindical/Web/ComercioLimeira/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do login
    while not _find_by_path('/html/body/div[2]/form[1]/div[1]/div[2]/table/tbody/tr[2]/td/input', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o usuario
    driver.find_element(by=By.ID, value='formLogin:i_login').send_keys(usuario)
    
    # insere a senha
    driver.find_element(by=By.ID, value='formLogin:i_senha').send_keys(senha)
    
    # clica em continuar
    driver.find_element(by=By.ID, value='formLogin:j_idt17').click()
    time.sleep(44)
    
    return driver, ''


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor = empresa[2]
    usuario = empresa[6]
    senha = empresa[7]
    
    print('3 - SINECOL - Sindicato dos Comerciários de Limeira')
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, usuario, senha)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {avisos}'
    
    return f'✔ Boleto gerado - {avisos}'
