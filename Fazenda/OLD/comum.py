# -*- coding: utf-8 -*-
from selenium.webdriver.support.ui import Select
from requests.packages import urllib3
from selenium import webdriver
from bs4 import BeautifulSoup
from requests import Session
from time import sleep
import OpenSSL.crypto, contextlib, tempfile, warnings, os, re
#import sys
#sys.path.append("..")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings('ignore')


def escreve_header_csv(texto):
    with open('Resumo.csv', 'r+') as f:
        conteudo = f.read()

    with open('Resumo.csv', 'w') as f:
        f.write(texto + '\n' + conteudo)

        
def atualiza_info(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    infos = (
        soup.find('input', attrs={'id': '__VIEWSTATE'}),
        soup.find('input', attrs={'id': '__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id': '__EVENTVALIDATION'})
    )

    lista = []
    for info in infos:
        try:
            lista.append(info.get('value'))
        except:
            lista.append('')

    return tuple(lista)


def escreve_relatorio_csv(texto):
    try:
        arquivo = open('resumo.csv', 'a')
    except:
        arquivo = open('resumo-error.csv', 'a')
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


def login_certificado(certificado, senha):
    print('>>> Logando como contabilista')

    url_login = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"
    with pfx_to_pem(certificado, senha) as cert:
        path_driver = os.path.join('..', 'phantomjs.exe')
        args = ['--ssl-client-certificate-file='+cert]

        driver = webdriver.PhantomJS(path_driver, service_args=args)
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()

        # Acessa a página inicial
        try:
            driver.get(url_login)
            sleep(1)
        except Exception as e:
            print(e)
            driver.quit()
            return False

        elems = [
            "ConteudoPagina_rdoListPerfil_1",
            "ConteudoPagina_btn_Login_Certificado_WebForms"
        ]
        for elem in elems:
            driver.find_element_by_id(elems).click(elem)
            sleep(1)

        cookies = driver.get_cookies()
        driver.quit()

    return cookies


def login_usuario(ident, username, senha, perfil):
    print('>>> Logando', ident)

    s = Session()
    url_login = f"https://cert01.fazenda.sp.gov.br/ca/ca"
    url_credetial = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"

    pagina_login = s.get(url_credetial)
    viewstate, viewstategenerator, eventvalidation = atualiza_info(pagina_login)
    info = {
        'ctl00$ScriptManager1': 'ctl00$ConteudoPagina$upnConsulta|ctl00$ConteudoPagina$btnAcessar',
        '__EVENTTARGET': '','__EVENTARGUMENT': '','__VIEWSTATE': viewstate,
        '__VIEWSTATEGENERATOR': viewstategenerator,'__VIEWSTATEENCRYPTED': '',
        '__EVENTVALIDATION': eventvalidation,'ctl00$ConteudoPagina$cboPerfil': str(perfil),
        'ctl00$ConteudoPagina$txtUsuario': username,'ctl00$ConteudoPagina$txtSenha': senha,
        '__ASYNCPOST': 'true','ctl00$ConteudoPagina$btnAcessar': 'Acessar'
    }
    s.post(url_credetial, info)

    perfil = 'CONTR' if perfil == 1 else 'CONTA'
    info = {
        'proxtela': '/deca.docs/contrib/servcontrvalid.htm',
        'SERVICO': perfil, 'ORIGEM': url_login, 'funcao': 'login',
        'UserId': username, 'PassWd': senha
    }

    pagina = s.post(url_login, info, verify=False)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    titulo = soup.find('title')
    
    if titulo.text == 'Controle de Acesso - Mensagem':
        return (False, f'Erro de login com usuario {username}')

    SID = re.search(r'SID=(\d+)', pagina.text)
    if not SID:
        return (False, 's_id não encontrado')

    return (s, SID.group(1))


def login_eCND(certificado, senha):
    print('>>> Logando como contabilista')
    
    url_base = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages'
    url_login = f'{url_base}/EmissaoCertidaoNegativa.aspx'
    url = f'{url_base}/Restrita/PesquisarContribuinte.aspx'
    
    with pfx_to_pem(certificado, senha) as cert:
        path_driver = os.path.join('..', 'phantomjs.exe')
        args = ['--ssl-client-certificate-file='+cert]

        driver = webdriver.PhantomJS(path_driver, service_args=args)
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()

        # Acessa a página inicial
        for i in range(5):
            try:
                driver.get(url)
                sleep(1)
                driver.find_element_by_id('btn_Login_Certificado_SSO_WebForms').click()
                sleep(1)
                break
            except Exception as e:
                if i == 4:
                    print(e.__class__)
                else:
                    print(" Nova Tentativa...")
        else:
            driver.quit()
            return False

        driver.get(url)
        cookies = driver.get_cookies()
        # driver.save_screenshot('debug_screen.png')
        driver.quit()

    return cookies