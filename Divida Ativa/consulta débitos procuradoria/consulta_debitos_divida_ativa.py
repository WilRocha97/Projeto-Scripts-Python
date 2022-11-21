# -*- coding: utf-8 -*-
import time, re, os, json
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def find_by_path(xpath, driver):
    try:
        elem = driver.find_element(by=By.XPATH, value=xpath)
        return elem
    except:
        return None
    
    
def login(empresa):
    settings = {"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}], "selectedDestinationId": "Save as PDF", "version": 2}
    prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings)}
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', prefs)
    options.add_argument('--kiosk-printing')
    
    status, driver = _initialize_chrome(options)
    cnpj, nome = empresa
    
    url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
    driver.get(url)

    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
    token = _solve_recaptcha(recaptcha_data)

    driver.find_element(by=By.ID, value='consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa').click()
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div/div[2]/div/div[3]/div/div[2]/div[2]/span/form/span/div/div[2]/table/tbody/tr/td[1]/div/div/span/select/option[2]').click()
    time.sleep(1)

    driver.find_element(by=By.ID, value='consultaDebitoForm:decTxtTipoConsulta:cnpj').send_keys(cnpj)
    driver.execute_script("document.getElementById('g-recaptcha-response').innerText='" + token + "'")
    time.sleep(1)

    driver.execute_script("document.getElementsByName('consultaDebitoForm:j_id102')[0].click()")
    time.sleep(1)

    resultado = consulta_debito(driver)
    return True


def consulta_debito(driver):
    situacao = re.compile(r'(consultaDebitoForm:dataTable:tb)').search(driver.page_source).group(1)
    
    if situacao:
        lista = re.findall(r'id=\"consultaDebitoForm:dataTable:(\d):lnkConsultaDebito', driver.page_source)
        
        for iten in lista:
            entrar_no_debito(driver, iten)
    else:
        return 'não tem débito'


def entrar_no_debito(driver, iten):
    driver.find_element(by=By.ID, value='consultaDebitoForm:dataTable:' + iten + ':lnkConsultaDebito').click()
    time.sleep(1)
    
    time.sleep(44)
    
    
@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        
        """erro = 'sim'
        while erro == 'sim':
            try:"""
        login(empresa)
        '''erro = 'nao'
        except:
            erro = 'sim'''


if __name__ == '__main__':
    run()
