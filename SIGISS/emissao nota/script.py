# -*- coding: utf-8 -*-
from win32com.client import Dispatch
from selenium import webdriver
from Dados import cadastros
import pyautogui as a
from time import sleep
import os


def get_chrome_version(file=None):
    paths = (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    )

    parser = Dispatch("Scripting.FileSystemObject")
    for file in paths:
        try:
            return parser.GetFileVersion(file)
        except Exception:
            continue
    return None


def download_chrome(version):
    import requests, wget, zipfile
    version = '.'.join(version.split('.')[:-1])

    try:
        os.remove('chromedriver.zip')
    except OSError:
        pass

    url = f'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{version}'
    response = requests.get(url)
    version = response.text

    download_url = f"https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip"
    latest_driver_zip = wget.download(download_url,'chromedriver.zip')

    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall()
    os.remove(latest_driver_zip)


def initialize_chrome():
    # get chromeBrowser version
    version = get_chrome_version()
    if version:
        try:
            # download chromeDriver version that matches
            download_chrome(version)
        except Exception as e:
            a.alert("Não foi possivel baixar o chromeDriver correpondente ao chromeBrowser")
            print(e)
            return False
    else:
        a.alert(r"Chrome browser não esta na pasta padrão C:\Program Files\Google\Chrome\Application")
        return False

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--ignore-certificate-errors")
    return webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)


def gera_nota(cnpj, valor, driver):
    url_pagina = "https://valinhos.sigissweb.com/nfecentral?oper=novo&MI=N"
    js_verifica_cnpj = "return document.getElementById('data_emissao').disabled"
    js_class = "document.getElementById('classif').value = '1.05A3.00'"
    js_desc = "document.getElementById('descricao').value = 'REFERENTE A UTILIZACAO DO SISTEMA DE EMISSÃO DE NOTAS'"
    js_emissao = "submitForm( document.forms[0] )"

    driver.get(url_pagina)
    sleep(1)
    # Clica no campo CNPJ
    a.click(170, 186)
    sleep(0.5)
    a.write(cnpj, interval=0.1)
    sleep(0.5)
    # Clica no Contribuinte
    a.click(286, 204)
    sleep(1.4)
    # Clica no campo Valor
    a.click(402, 251)
    sleep(0.5)
    a.write(valor, interval=0.1)
    a.press('tab')
    sleep(0.5)
    if driver.execute_script(js_verifica_cnpj):
        print(f'>>> Empresa {cnpj} não cadastrada')
        return False
    driver.execute_script(js_class)
    sleep(0.3)
    driver.execute_script(js_desc)
    sleep(0.7)
    driver.execute_script(js_emissao)
    sleep(1)


def check_for_duplicates(raw_lista):
    lista = []

    for i in raw_lista:
        if i not in lista:
            lista.append(i)
        else:
            print("Empresa duplicada", i)

    return len(lista) == len(raw_lista)


def emite_nota():
    url_login = 'https://valinhos.sigissweb.com'
    js_login = "document.getElementById('edtlogin').value = '24982859000101'"
    js_senha = "document.getElementById('edtsenha').value = '58269953'"
    js_acesso = "submitForm( document.forms[0] );"
    
    driver = initialize_chrome()
    if not driver: return False
    
    driver.get(url_login)
    sleep(1)
    driver.execute_script(js_login)
    driver.execute_script(js_senha)
    sleep(1)
    driver.execute_script(js_acesso)
    sleep(2)

    # Verificar por empresa duplicadas
    raw_lista = [item[0] for item in cadastros]
    if not check_for_duplicates(raw_lista):
        print("Processo encerrado por conflito de dados")
        driver.quit()
        return  False

    for index, cadastro in enumerate(cadastros, 1):
        print(f'Emitido para empresa {cadastro[0]} - {index}')
        try:
            gera_nota(*cadastro, driver)
            sleep(2)
        except Exception as e:
            print(e)

    driver.quit()
    
    
if __name__ == '__main__':
    emite_nota()
