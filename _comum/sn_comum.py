# -*- coding: utf-8 -*-
import time

import timer
from bs4 import BeautifulSoup
from requests import Session
from time import sleep
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, re
import warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from captcha_comum import _solve_text_captcha, _solve_hcaptcha
from selenium_comum import send_input
from chrome_comum import initialize_chrome


# variaveis globais
_servs = {
    'snparc': '37',
    'parcmei': '48',
    'pertsn': '61',
    'recibos': '60',
    'caixa postal': '41',
}


def find_by_id(id_nome, driver):
    try:
        elem = driver.find_element(by=By.ID, value=id_nome)
        return elem
    except:
        return None


def new_session_sn(cnpj, cpf, cod, serv, driver, usa_driver=False):
    url_base = 'http://www8.receita.fazenda.gov.br/SimplesNacional'
    url_login = f'{url_base}/controleAcesso/Autentica.aspx?id={_servs[serv]}'
    aux_acessos = zip(('txtCNPJ', 'txtCPFResponsavel', 'txtCodigoAcesso'), (cnpj, cpf, cod))
    
    for url in (url_base, url_login):
        driver.get(url)
        sleep(1)
        
    print('>>> Aguardando site')
    timer = 0
    while not find_by_id('captcha-img', driver):
        sleep(1)
        timer += 1
        if timer >= 60:
            new_session_sn(cnpj, cpf, cod, serv, driver, usa_driver=False)
        
    element = driver.find_element(by=By.ID, value='captcha-img')
    location = element.location
    size = element.size
    driver.save_screenshot('ignore\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    sleep(2)
    im = Image.open(r'ignore\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'ignore\captcha\captcha.png')
    sleep(1)

    captcha = _solve_text_captcha(os.path.join('ignore', 'captcha', 'captcha.png'))

    send_input('txtTexto_captcha_serpro_gov_br', captcha, driver)

    for key, value in aux_acessos:
        send_input(f'ctl00_ContentPlaceHolder_{key}', value, driver)
        sleep(0.5)
    
    print('>>> Logando na empresa')
    elem = driver.find_element(by=By.ID, value='ctl00_ContentPlaceHolder_btContinuar')
    elem.click()
    sleep(1)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    erro = soup.find('span', attrs={'id': 'ctl00_ContentPlaceHolder_lblErro'})
    if erro:
        driver.get(url_base)
        if usa_driver:
            return driver, f'Erro Login - {erro.text.strip()}'
        else:
            return f'Erro Login - {erro.text.strip()}'

    session = Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])
    
    if usa_driver:
        return driver, session
    else:
        driver.delete_all_cookies()
        return session
_new_session_sn = new_session_sn


# não funciona porque o site reconhece robô
def new_session_sn_driver(cnpj):
    url_home = "http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao"
    _site_key = '2c0f2c5b-d8b9-469a-98ec-56271c2f68e4'

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    status, driver = initialize_chrome()
    driver.get(url_home)

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    padrao = re.compile(r'id=\"(h-captcha-response-.+)\" name')
    resposta = padrao.search(str(soup))
    id_captcha = resposta.group(1)
    print(id_captcha)

    # gera o token para passar pelo captcha
    hcaptcha_data = {'sitekey': _site_key, 'url': url_home}
    token = _solve_hcaptcha(hcaptcha_data)
    token = str(token)

    print(f'>>> Logando na empresa {cnpj}')
    element = driver.find_element(by=By.ID, value='cnpj')
    element.send_keys(cnpj)

    script = 'document.getElementById("{}").innerHTML="{}";'.format(id_captcha, token)
    driver.execute_script(script)
_new_session_sn_driver = new_session_sn_driver
