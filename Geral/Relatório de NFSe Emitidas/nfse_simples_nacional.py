# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, usuario, senha):
    print('>>> Logando no site')
    try:
        driver.get('https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional')
    except:
        print('>>> Site demorou pra responder, tentando novamente')
        return driver, 'erro'
    time.sleep(1)
    
    _send_input('Inscricao', usuario, driver)
    _send_input('Senha', senha, driver)
    time.sleep(1)
    
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button').click()
    time.sleep(1)
    return driver, 'ok'
    
def consulta_notas(cnpj, nome, driver):
    print('>>> Consultando notas emitidas')
    driver.get('https://www.nfse.gov.br/EmissorNacional/Notas/Emitidas')
    time.sleep(1)
    
    notas = re.compile(r'<td class=\"td-data\">\n\s+(\d\d/\d\d/\d\d\d\d)(\n.+){6}'
                      r'<span class=\"cnpj\">(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)'
                      r'.+\n.+\n(.+)( ){20}</div>(\n.+){2}\s+(.+)(\n.+){2}\s+(.+)'
                      r'(\n.+){3}.+img/tb-(.+).svg.+(\n.+){2}data-original-title=\"(.+)\"').findall(driver.page_source)
    
    if not notas:
        print(driver.page_source)
    
    print('>>> Capturando dados das notas emitidas')
    for nota in notas:
        geracao = nota[0]
        emitida_para = f'{nota[2]};{nota[3]}'
        municipio = nota[6].replace('/', ';')
        valor = nota[8]
        situacao = nota[10]
        impostos = nota[12]
        
        dados = f'{cnpj};{nome};{geracao};{emitida_para};{municipio};{valor};{situacao};{impostos}'
        _escreve_relatorio_csv(dados)
        print(dados.replace(';', ' - '))
        
    return driver


@_time_execution
def run():
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
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, usuario, senha = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            driver, situacao = login(driver, usuario, senha)
            if situacao == 'ok':
                driver = consulta_notas(cnpj, nome, driver)
                break
            driver.close()
        driver.close()
        
if __name__ == '__main__':
    run()
