# -*- coding: utf-8 -*-
from warnings import filterwarnings
from bs4 import BeautifulSoup
from urllib3 import exceptions, disable_warnings
from requests import Session
from captcha_comum import _solve_recaptcha
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import time, re, os, warnings

from comum_comum import content_or_soup, _pfx_to_pem
from chrome_comum import initialize_chrome

disable_warnings(exceptions.InsecureRequestWarning)
warnings.filterwarnings('ignore')

# variáveis globais
_perfil = {'contribuinte': ('1', 'CONTR'), 'contador': ('2', 'CONTA')}


def salvar_arquivo(nome, dados):
    try:
        arquivo = open(os.path.join('execução/documentos', nome), 'wb')
    except FileNotFoundError:
        os.makedirs('documentos', exist_ok=True)
        arquivo = open(os.path.join('documentos', nome), 'wb')

    for parte in dados.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()
    
    
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


_atualiza_info = atualiza_info


@content_or_soup
def get_info_post(soup):
    infos = [
        soup.find('input', attrs={'id': '__VIEWSTATE'}),
        soup.find('input', attrs={'id': '__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id': '__EVENTVALIDATION'})
    ]
    
    # state, generator, validation
    return tuple(info.get('value', '') for info in infos if info)


_get_info_post = get_info_post


# Loga no site da sefaz/login como contribuinte ou contador usando login e senha
# Retorna tupla com instância de Session com os cookies
# da sessão e o id da sessão em caso de sucesso
# Retorna string mensagem em caso de erro
def new_session_fazenda(ni, user, pwd, tipo):
    print('\n>>> logando', ni)
    filterwarnings('ignore')
    
    base = 'ctl00$ConteudoPagina'
    url_login = "https://cert01.fazenda.sp.gov.br/ca/ca"
    url_home = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"
    _site_key = '6LesbbcZAAAAADrEtLsDUDO512eIVMtXNU_mVmUr'
    
    session = Session()
    res = session.get(url_home, verify=False)
    state, generator, validation = get_info_post(content=res.content)
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': _site_key, 'url': url_home}
    token = _solve_recaptcha(recaptcha_data)
    
    data = {
        'ctl00$ScriptManager1': f'{base}$upnConsulta|{base}$btnAcessar',
        '__EVENTTARGET': '',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': state,
        '__VIEWSTATEGENERATOR': generator,
        '__VIEWSTATEENCRYPTED': '',
        '__EVENTVALIDATION': validation,
        f'{base}$cboPerfil': _perfil[tipo][0],
        f'{base}$txtUsuario': user,
        f'{base}$txtSenha': pwd,
        'g-recaptcha-response': token,
        '__ASYNCPOST': 'true',
        f'{base}$btnAcessar': 'Acessar'
    }
    
    pagina = session.post(url_home, data, verify=False)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    print(soup)
    
    data = {
        'proxtela': '/deca.docs/contrib/servcontrvalid.htm',
        'SERVICO': _perfil[tipo][1], 'ORIGEM': url_login,
        'funcao': 'login', 'UserId': user, 'PassWd': pwd
    }
    
    res = session.post(url_login, data, verify=False)
    if res.url == url_login:
        return f'Erro login com usuario {user}'
    
    return session, res.url[69:98]


_new_session_fazenda = new_session_fazenda


def new_session_fazenda_driver(user, pwd, perfil, retorna_driver=False, options='padrão'):
    url_home = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"
    _site_key = '6LesbbcZAAAAADrEtLsDUDO512eIVMtXNU_mVmUr'

    if options == 'padrão':
        # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--window-size=1920,1080')
        # options.add_argument("--start-maximized")
    
    status, driver = initialize_chrome(options)
    driver.get(url_home)
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': _site_key, 'url': url_home}
    token = _solve_recaptcha(recaptcha_data)
    token = str(token)
    
    if perfil == 'contador':
        button = driver.find_element(by=By.ID, value='ConteudoPagina_rdoListPerfil_1')
        button.click()
        time.sleep(1)
    
    elif perfil != 'contribuinte':
        driver.close()
        return False
    
    print(f'>>> Logando no usuário')
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtUsuario')
    element.send_keys(user)
    
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtSenha')
    element.send_keys(pwd)
    
    script = 'document.getElementById("g-recaptcha-response").innerHTML="{}";'.format(token)
    driver.execute_script(script)
    
    script = 'document.getElementById("ConteudoPagina_btnAcessar").disabled = false;'
    driver.execute_script(script)
    time.sleep(1)
    
    button = driver.find_element(by=By.ID, value='ConteudoPagina_btnAcessar')
    button.click()
    time.sleep(3)
    
    try:
        button = driver.find_element(by=By.XPATH, value='/html/body/div[2]/section/div/div/div/div[2]/div/ul/li/form/div[5]/div/a')
        button.click()
        time.sleep(3)
    except:
        pass
    
    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    padrao = re.compile(r'SID=(.\d+)')
    resposta = padrao.search(str(soup))
    
    if not resposta:
        try:
            padrao = re.compile(r'(Senha inserida está incorreta)')
            resposta = padrao.search(str(soup))
            
            if not resposta:
                padrao = re.compile(r'(Favor informar o login e a senha corretamente.)')
                resposta = padrao.search(str(soup))
                
                if not resposta:
                    padrao = re.compile(r'(O usuário não tem acesso a este serviço do sistema ou o serviço não foi liberado para a sua utilização.)')
                    resposta = padrao.search(str(soup))
                    
                    if not resposta:
                        padrao = re.compile(r'(ERRO INTERNO AO SISTEMA DE CONTROLE DE ACESSO)')
                        driver.save_screenshot(r'ignore\debug_screen.png')
                        resposta = padrao.search(str(soup))
            
            sid = resposta.group(1)
            print(f'❌ {sid}')
            cokkies = 'erro'
            driver.close()
            
            return cokkies, sid
        except:
            driver.save_screenshot(r'ignore\debug_screen.png')
            driver.close()
            return False
    
    if retorna_driver:
        sid = resposta.group(1)
        return driver, sid
    
    sid = resposta.group(1)
    cookies = driver.get_cookies()
    driver.quit()
    
    return cookies, sid
_new_session_fazenda_driver = new_session_fazenda_driver


def login_ecnd(certificado, senha):
    certificados = {'ADELINO': r'..\..\_cert\CERT ADELINO 45274040.pfx',
                    'RPEM': r'..\..\_cert\CERT RPEM 35586086.pfx',
                    'RODRIGO': r''}
    
    print('>>> Logando como contabilista')
    
    url_base = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages'
    # url_login = f'{url_base}/EmissaoCertidaoNegativa.aspx'
    url = f'{url_base}/Restrita/PesquisarContribuinte.aspx'
    
    with _pfx_to_pem(certificados[certificado], senha) as cert:
        path_driver = os.path.join('..', 'phantomjs.exe')
        args = ['--ssl-client-certificate-file=' + cert]
        
        driver = webdriver.PhantomJS(path_driver, service_args=args)
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()
        
        # Acessa a página inicial
        for i in range(5):
            try:
                driver.get(url)
                sleep(1)
                driver.find_element(by=By.ID, value='btn_Login_Certificado_SSO_WebForms').click()
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


_login_ecnd = login_ecnd
