# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, shutil


from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome
from pyautogui_comum import _click_img, _wait_img, _find_img


def localiza_elemento(driver, elemento):
    try:
        driver.find_element(by=By.XPATH, value=elemento)
        return True
    except:
        return False


def renomear(empresa, apuracao, vencimento):
    cnpj, nome, valor, cod = empresa
    download_folder = "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execucao\\Guias"
    guia = os.path.join(download_folder, 'Darf.pdf')
    while not os.path.exists(guia):
        time.sleep(1)
    while os.path.exists(guia):
        try:
            arquivo = f'{nome.replace("/", " ")} - {cnpj} - DARF IRRF {cod} {apuracao.replace("/", "-")} - venc. {vencimento.replace("/", "-")}.pdf'
            shutil.move(guia, os.path.join(download_folder, arquivo))
            time.sleep(2)
        except:
            pass


def login_veri(empresa, driver):
    driver.get('https://26973312000175.veri-sp.com.br/login')
    print('>>> Acessando o site')
    time.sleep(1)

    # inserir o CNPJ no campo
    driver.find_element(by=By.NAME, value='login').send_keys('joao@veigaepostal.com.br')
    time.sleep(1)

    # inserir a senha no campo
    driver.find_element(by=By.NAME, value='senha').send_keys('Milenio03')
    time.sleep(1)

    # clica em acessar
    driver.find_element(by=By.NAME, value='btn_login').click()
    time.sleep(1)
    
    # gerar a guia de DCTF
    gerar(empresa, driver)


def gerar(empresa, driver):
    cnpj, nome = empresa
    while not driver.find_element(by=By.ID, value='carouselExampleIndicators'):
        time.sleep(1)
        
    driver.get('https://26973312000175.veri-sp.com.br/dctf_web_nova/index?filter=ATIVA')
    time.sleep(1)
    
    while not driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input'):
        time.sleep(1)

    driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input').send_keys(cnpj)
    time.sleep(2)
    
    while driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[2]/div/div/div[2]/table/tbody/tr[2]/td[11]/div/div/a'):
        driver.execute_script("document.querySelector('#example > tbody > tr.odd > td:nth-child(11) > div > div > a').click()")
        time.sleep(2)
        
        while localiza_elemento(driver, '/html/body/div[6]/div/h2'):
            print('Gerando...')
            time.sleep(1)
    
        while not localiza_elemento(driver, '/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input'):
            time.sleep(1)

        driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input').send_keys(cnpj)
        time.sleep(2)

    
@_time_execution
def run():
    os.makedirs('execucao/Guias', exist_ok=True)
    # p.mouseInfo()

    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execucao\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)

        # iniciar o driver do chome
        status, driver = _initialize_chrome(options)

        # fazer login do SICALC
        login_veri(empresa, driver)
        
        driver.close()


if __name__ == '__main__':
    run()
