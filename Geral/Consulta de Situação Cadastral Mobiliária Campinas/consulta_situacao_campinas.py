# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, re, fitz, shutil

from sys import path

path.append(r'..\..\_comum')
from captcha_comum import _solve_text_captcha
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, cnpj):
    print('>>> Logando no site')
    timer = 0
    while not _find_by_id('captcha_image', driver):
        try:
            driver.get('https://situacao.campinas.sp.gov.br/')
        except:
            print('>>> Site demorou pra responder, tentando novamente')
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'
    
    
    # resolver o captcha
    captcha = _solve_text_captcha(driver, 'captcha_image')
    
    driver.find_element(by=By.ID, value='radio3').click()
    _send_input('xInscricao', cnpj, driver)
    _send_input('cap_text', captcha, driver)
    driver.find_element(by=By.XPATH, value='/html/body/form/table[2]/tbody/tr[3]/td/table[2]/tbody/tr[1]/td[1]/input').click()
    time.sleep(3)
    print(driver.page_source)
    time.sleep(33)
    
    return driver, 'ok'


def consulta_notas(download_folder, driver, cnpj, nome):
    print('>>> Consultando notas emitidas')
    
    
    return driver




@_time_execution
def run():
    download_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\Consulta de Situação Cadastral Mobiliária Campinas\\Execução"
    os.makedirs(download_folder, exist_ok=True)
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,  # muda o diretório padrão de download do navegador
        "download.prompt_for_download": False,  # faz o download automatico sem perguntar onde salvar
        "download.directory_upgrade": True,  # atualiza o diretório de download padrão do navegador
        "plugins.always_open_pdf_externally": True  # não irá abrir o PDF no navegador
    })
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            driver, situacao = login(driver, cnpj)
            if situacao == 'ok':
                driver = consulta_notas(download_folder, driver, cnpj, nome)
                break
            
            driver.close()
        driver.close()
    

if __name__ == '__main__':
    run()
