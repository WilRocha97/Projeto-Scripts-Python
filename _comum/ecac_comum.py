# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from urllib import parse
from time import sleep
import re

from selenium_comum import exec_phantomjs
from captcha_comum import break_hcaptcha
from comum_comum import pfx_to_pem, _headers


# variaveis globais
_site_key = '903db64c-2422-4230-a22e-5645634d893f'


# Loga no site ecac com o certificado atravez de um webdriver, cria 
# uma instancia de Session e passa os cookies do webdriver para a Session
# Retorna uma instancia de session com os cookies em caso de sucesso
# Retorna uma mensagem de erro em caso de erro
def new_session_ecac(cert, pwd, timeout=20, delay=1):
    url = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    print('\n>>> nova sessão com', cert.split('\\')[-1])

    with pfx_to_pem(cert, pwd) as cert:
        driver = exec_phantomjs(cert=cert, headers=_headers)
        if isinstance(driver, str): return driver

    driver.get(url)
    recaptcha_data = {'sitekey': _site_key, 'url': url}
    token = break_hcaptcha(recaptcha_data)

    if isinstance(token, str):
        driver.quit()
        return token

    print('>>> enviando form')
    script = f"document.getElementsByName('h-captcha-response')[1].innerText = '{token['code']}';"
    driver.execute_script(script)

    elem = driver.find_element_by_id('frmLoginCert')
    elem.submit()

    for i in range(timeout):
        sleep(delay)
        try:
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            url_main = soup.find('div', attrs={'id': 'cert-digital'}).a.get('href', '')
            if not url_main: return 'Não encontrou url principal'
            break
        except AttributeError:
            continue
    else:
        driver.save_screenshot(r'ignore\debug_screen.png')
        return 'Limite de tentativas de recuperar url principal atingidas'

    driver.get(url_main)
    sleep(1)

    session = Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])

    driver.quit()
    return session


_new_session_ecac = new_session_ecac


# Troca de cnpj no site ecac depois de logado
# Retorna um dicionario com {key: bool, value: str}
# onde key é True se logou e False se não logou e value é uma mensagem 
def new_login_ecac(cnpj, session, limit=5):
    cnpj = ''.join(i for i in cnpj if i.isdigit()).rjust(14, '0')
    url = 'https://cav.receita.fazenda.gov.br/autenticacao/api/mudarpapel/procuradorpj'
    url_captcha = 'https://cav.receita.fazenda.gov.br/ecac/Default.aspx'

    recaptcha_data = {'sitekey': _site_key, 'url': url_captcha}
    token = break_hcaptcha(recaptcha_data)

    if isinstance(token, str):
        return { 'key': False, 'value': token }

    res = None
    data = {'ni': cnpj, 'hCaptchaResponse': token['code']}

    for i in range(limit):
        try:
            response = session.post(url, data, headers=_headers, verify=False)
            soup = BeautifulSoup(response.content, 'html.parser')
            key = soup.find('key').get_text() == 'true'

            res = { 'key': key, 'value': soup.find('value').get_text().strip().lower()}
            if key:
                break
        except AttributeError:
            continue
    else:
        if res is None:
            res = {'key': False, 'value': 'limite de tentativas atingido'}

    return res
_new_login_ecac = new_login_ecac
