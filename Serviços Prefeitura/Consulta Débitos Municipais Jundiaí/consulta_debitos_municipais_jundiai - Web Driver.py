# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from requests import Session
from PIL import Image
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha

os.makedirs('execucao/Certidões', exist_ok=True)
site_key = '6LfK0igTAAAAAOeUqc7uHQpW4XS3EqxwOUCHaSSi'


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


def pesquisar(cnpj, insc_muni):
    status, driver = initialize_chrome()
    
    url_inicio = 'https://jundiai.sp.gov.br/servicos-online/certidao-negativa-de-debitos-mobiliarios/'
    driver.get(url_inicio)
    
    while not find_by_id('g-recaptcha-response', driver):
        time.sleep(1)

    data = {'url': url_inicio, 'sitekey': site_key}
    response = _solve_recaptcha(data)
    
    _send_input('DadoContribuinteMobiliario1_txtCfm', insc_muni, driver)
    _send_input('DadoContribuinteMobiliario1_txtNrCicCgc', cnpj, driver)
    _send_input('g-recaptcha-response', response, driver)
    
    
@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, insc_muni = empresa

        _indice(count, total_empresas, empresa)

        situacao, situacao_print = login(cnpj, insc_muni)
        _escreve_relatorio_csv(f'{cnpj};{insc_muni};{situacao}')
        print(situacao_print)


if __name__ == '__main__':
    run()
