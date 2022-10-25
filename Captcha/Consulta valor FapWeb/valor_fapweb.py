# -*- coding: utf-8 -*-
import re
from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _download_file, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome, _find_by_id

_site_key = '6LcUdlwaAAAAADjMUnV5jwUKKMlqdZbKNKXTfYEi'
_execucao = 'Valor FAP original'


def login(cnpj_sem_mil_contra, cnpj, driver):
    tentativas = 1
    x = 'erro'
    while x == 'erro':
        driver.get('https://google.com')
        url = 'https://www2.dataprev.gov.br/FapWeb/pages/login.xhtml'
        print('>>> Acessando o site')
        driver.get(url)

        recaptcha_data = {'sitekey': _site_key, 'url': url}

        captcha = _solve_recaptcha(recaptcha_data)
        print('>>> logando no site')

        try:
            driver.find_element(by=By.ID, value='form:CNPJ').send_keys(cnpj_sem_mil_contra)
            driver.find_element(by=By.ID, value='form:senha').send_keys(cnpj_sem_mil_contra)
            driver.execute_script('document.getElementById("g-recaptcha-response").innerHTML="' + captcha + '";')
    
            sleep(1)
            driver.find_element(by=By.ID, value='form:btnAutenticar').click()
            sleep(5)
    
            html = driver.page_source.encode('utf-8')
            soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
            padrao = re.compile(r'<div class=\"mensagem\"><ul><li class=\"erro\">(.+)</li></ul></div>')
            resposta = padrao.search(str(soup))
            if resposta:
                resposta = resposta.group(1)
    
                print('❌ {}'.format(resposta))
                tentativas += 1
                x = 'erro'
            else:
                x = 'ok'
        except:
            resposta = 'Erro no login'
            print('❌ {}'.format(resposta))
            tentativas += 1
            x = 'erro'

        if tentativas >= 3:
            _escreve_relatorio_csv(';'.join([cnpj, resposta]), nome=_execucao)
            return False

    if not consulta(driver, cnpj):
        return False

    return True


def consulta(driver, cnpj):
    print('>>> consultando')
    driver.get('https://www2.dataprev.gov.br/FapWeb/pages/consulta/resultadoConsultaFap.xhtml')

    while not _find_by_id('form:panelConsulta', driver):
        sleep(1)

    sleep(1)

    # selecionar uma competência
    # driver.execute_script('document.getElementById("form:anoVigencia").value=2020')
    # driver.execute_script('document.getElementById("form:btnConsultaCnpjDigitado").click(2)')
    # sleep(5)

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')

    try:
        padrao = re.compile(r'(Valor do FAP Original)(\n.+){8}\">(.+)(\n.+){5}\">(.+)')
        resposta = padrao.search(str(soup))
        valor = resposta.group(1) + ': ' + resposta.group(3)
        data = resposta.group(5)

        print('✔ {} - {}'.format(valor, data))
        _escreve_relatorio_csv(';'.join([cnpj, valor, data]), nome=_execucao)

    except:
        try:
            padrao = re.compile(r'<div class=\"mensagem\"><ul><li class=\"erro\">(.+).</li><li')
            resposta = padrao.search(str(soup))
            erro = resposta.group(1)

            print('❌ {}'.format(erro))
            _escreve_relatorio_csv(';'.join([cnpj, erro]), nome=_execucao)
        except:
            print('❌ Valor do FAP Original não encontrado')
            _escreve_relatorio_csv(';'.join([cnpj, 'Valor do FAP Original não encontrado']), nome=_execucao)
            return False

    return True


@_time_execution
def run():
    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cnpj_sem_mil_contra = empresa

        _indice(count, total_empresas, empresa)

        status, driver = _initialize_chrome(options)

        if not login(cnpj_sem_mil_contra, cnpj, driver):
            continue

        driver.close()


if __name__ == '__main__':
    run()
