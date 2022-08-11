# -*- coding: utf-8 -*-
from requests.packages import urllib3
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from functools import wraps
from urllib import parse
from time import sleep
import OpenSSL.crypto
import contextlib
import tempfile
import warnings
import os

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings('ignore')

# decorator
def time_execution(func):
    @wraps(func)
    def wrapper():
        comeco = datetime.now()
        print("Execução iniciada as:", comeco)
        func()
        print("Tempo de execução:", datetime.now() - comeco)

    return wrapper


def escreve_header_csv(texto):
    with open('Resumo.csv', 'r+') as f:
        conteudo = f.read()

    with open('Resumo.csv', 'w') as f:
        f.write(texto + '\n' + conteudo)


def escreve_relatorio_csv(texto):
    try:
        arquivo = open("Resumo.csv", 'a')
    except:
        arquivo = open("Resumo-error.csv", 'a')
    arquivo.write(texto+'\n')
    arquivo.close()


def registra_ocorrencia(cnpj, texto):
    nome = os.path.join('documentos', cnpj+' - ocorrencias.txt')
    try:
        arquivo = open(nome, 'a')
    except FileNotFoundError:
        os.makedirs('documentos', exist_ok=True)
        arquivo = open(nome, 'w')
    arquivo.write(texto+'\n')
    arquivo.close()


def salvar_arquivo(nome, dados):
    try:
        arquivo = open(os.path.join('documentos', nome), 'wb')
    except FileNotFoundError:
        os.makedirs('documentos', exist_ok=True)
        arquivo = open(os.path.join('documentos', nome), 'wb')

    for parte in dados.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()


def atualiza_info(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')

    infos = [
        soup.find('input', attrs={'id':'__VIEWSTATE'}),
        soup.find('input', attrs={'id':'__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id':'__EVENTVALIDATION'}),
        soup.find('input', attrs={'id':'mensagemComBase64'}),
    ]

    result = []
    for info in infos:
        try:
            result.append(info.get('value'))
        except:
            result.append('')

    return tuple(result) # state, generator, validation, msg64


@contextlib.contextmanager
def pfx_to_pem(pfx_path, pfx_password):
    ''' Decrypts the .pfx file to be used with requests. '''
    with tempfile.NamedTemporaryFile(suffix='.pem', delete=False) as t_pem:
        f_pem = open(t_pem.name, 'wb')
        pfx = open(pfx_path, 'rb').read() # --> type(pfx) > bytes
        p12 = OpenSSL.crypto.load_pkcs12(pfx, pfx_password)
        data_inicial = p12.get_certificate().get_notBefore()
        data_vencimento = p12.get_certificate().get_notAfter()
        if p12.get_certificate().has_expired(): 
            print('Certificado possivelmente vencido')
            data_inicial = p12.get_certificate().get_notBefore()
            data_vencimento = p12.get_certificate().get_notAfter()
            print(f'Inicio: {data_inicial[6:8]}/{data_inicial[4:6]}/{data_inicial[:4]}')
            print(f'Vencimento: {data_vencimento[6:8]}/{data_vencimento[4:6]}/{data_vencimento[:4]}')
        f_pem.write(OpenSSL.crypto.dump_privatekey(OpenSSL.crypto.FILETYPE_PEM, p12.get_privatekey()))
        f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, p12.get_certificate()))
        ca = p12.get_ca_certificates()
        if ca is not None:
            for cert in ca:
                f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, cert))
        f_pem.close()

        yield t_pem.name


def new_session(certificado, senha):
    print('>>> Nova Sessão', certificado)
    url_main = 'https://cav.receita.fazenda.gov.br/ecac/'
    url_aux = 'https://cav.receita.fazenda.gov.br/autenticacao/login'

    headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

    for key, value in enumerate(headers):
        capability_key = 'phantomjs.page.customHeaders.{}'.format(key)
        webdriver.DesiredCapabilities.PHANTOMJS[capability_key] = value

    with pfx_to_pem(certificado, senha) as cert:
        driver_path = os.path.join('..', 'phantomjs.exe')
        args = ['--ssl-client-certificate-file='+cert, '--ignore-ssl-errors=true']

        driver = webdriver.PhantomJS(driver_path, service_args=args)
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()
        driver.get(url_aux)
        sleep(1)
        # Acessa a página inicial

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    base_state = soup.find('div', attrs={'id':'login-dados-certificado'}).a.get('href')
    state = base_state.split('=')[-1]

    query = {
        'response_type': 'code', 'client_id': 'cav.receita.fazenda.gov.br',
        'scope': 'openid+govbr_recupera_certificadox509+govbr_confiabilidades',
        'redirect_uri': 'https://cav.receita.fazenda.gov.br/autenticacao/login/govbrsso',
        'state': state
    }

    url_base = 'https://certificado.sso.acesso.gov.br/authorize'
    parsed_url = parse.urlparse(url_base)._replace(query=parse.urlencode(query, safe='+'))
    url_login = parse.urlunparse(parsed_url)

    while True:
        try:
            driver.get(url_login)
            sleep(1)
        except Exception as e:
            print(e)
            driver.quit()
            return False

        if driver.current_url != url_main:
            print('\t--nova tentativa', driver.current_url)
            continue

        session = Session()
        for cookie in driver.get_cookies():
            session.cookies.set(cookie['name'], cookie['value'])
        break

    driver.quit()
    return session


def new_login(cnpj, session):
    print('>>> Logando', cnpj)
    url = 'https://cav.receita.fazenda.gov.br/autenticacao/api/mudarpapel/procuradorpj'

    headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

    response = session.post(url, {"" : cnpj}, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')
    data = { 
        'Key': soup.find('key').get_text() == 'true',
        'Value': soup.find('value').get_text()
    }

    return data


def new_login_credential(ni, user, senha):
    print('>>> Logando', ni)

    url_base = 'https://cav.receita.fazenda.gov.br'
    url_acesso = f'{url_base}/autenticacao/Login/CodigoAcesso'
    url_home = f'{url_base}/ecac/'

    info = {
        'ExibeCaptcha': 'False', 'id': '-1', 'NI': ni,
        'CodigoAcesso': user, 'Senha': senha, 'ExibiuImagem': 'true',
    }

    session = Session()
    response = session.post(url_acesso, info, verify=False)
    soup = BeautifulSoup(response.content, 'html.parser')
    msg = soup.find('div', attrs={'class':'login-caixa-erros-validacao'})
    
    if not msg:
        session.get(url_home)
        return True, session

    msg = msg.text.replace('Atenção: ', '').replace('Atenção:', '')
    return False, msg.strip()
    