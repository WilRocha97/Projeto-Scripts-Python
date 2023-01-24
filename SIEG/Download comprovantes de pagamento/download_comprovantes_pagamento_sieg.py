# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import prompt
import os, time, re, csv

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome


def localiza_path(driver, elemento):
    try:
        driver.find_element(by=By.XPATH, value=elemento)
        return True
    except:
        return False


def localiza_id(driver, elemento):
    try:
        driver.find_element(by=By.ID, value=elemento)
        return True
    except:
        return False
    

def login_sieg(driver):
    driver.get('https://auth.sieg.com/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # inserir o CNPJ no campo
    driver.find_element(by=By.ID, value='txtEmail').send_keys('willian.rocha@veigaepostal.com.br')
    time.sleep(1)
    
    # inserir a senha no campo
    driver.find_element(by=By.ID, value='txtPassword').send_keys('Milenniumfalcon9')
    time.sleep(1)
    
    # clica em acessar
    driver.find_element(by=By.ID, value='btnSubmit').click()
    time.sleep(1)
    
    return driver
    

def sieg_iris(driver):
    print('>>> Acessando IriS DCTF WEB')
    driver.get('https://hub.sieg.com/IriS/#/ComprovantesDePagamentos')

    return driver


def procura_empresa(competencia, empresa, driver, options):
    cnpj, nome = empresa
    # espera a barra de pesquisa, se não aparecer em 1 minuto, recarrega a página
    timer = 0
    while not localiza_id(driver, 'select2-ddlCompanyIris-container'):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            print('>>> Teantando novamente\n')
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)
            driver = sieg_iris(driver)
            timer = 0
        
    time.sleep(5)
    print('>>> Pesquisando empresa')
    
    # under construction
    
    return driver


@_time_execution
def run():
    competencia = prompt(text='Qual competência referênte?', title='Script incrível', default='00/0000')
    os.makedirs('execução/Guias', exist_ok=True)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    contador = 1
    while 1 < 2:
        print(f'\n\nIniciando rotina Nº {contador} ----------------------------------------------------------\n')
        # abre a planilha de dados
        file = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Dados.csv"
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()

        empresas = list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))

        # configurar um indice para a planilha de dados
        if contador == 1:
            index = _where_to_start(tuple(i[0] for i in empresas))
            if index is None:
                return False
        
        else:
            index = 0
        
        # iniciar o driver do chome
        status, driver = _initialize_chrome(options)
        driver = login_sieg(driver)
        
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
    
            # configurar o indice para localizar em qual empresa está
            _indice(count, total_empresas, empresa)
            erro = 'sim'
            while erro == 'sim':
                try:
                    driver = sieg_iris(driver)
                    driver = procura_empresa(competencia, empresa, driver, options)
                    erro = 'não'
                except:
                    try:
                        erro = 'sim'
                        driver.close()
                        status, driver = _initialize_chrome(options)
                        driver = login_sieg(driver)
                    except:
                        erro = 'sim'
                        status, driver = _initialize_chrome(options)
                        driver = login_sieg(driver)

        contador += 1
        os.remove("V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\resumo.csv")
        

if __name__ == '__main__':
    run()
