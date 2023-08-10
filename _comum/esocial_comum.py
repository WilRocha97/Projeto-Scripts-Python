# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from time import sleep

from selenium_comum import exec_phantomjs
from comum_comum import pfx_to_pem, _headers

# variáveis globais
_allowed_urls = (
    'https://www.esocial.gov.br/portal/Home/Inicial?tipoEmpregador=EMPREGADOR_GERAL',
    'https://www.esocial.gov.br/portal/Home/Inicial?tipoEmpregador=EMPREGADOR_DOMESTICO_GERAL'
)


# Loga no site esocial com o certificado através do webdriver, cria
# uma instância de Session e passa os cookies do webdriver para a Session
# Retorna uma instância de session com os cookies em caso de sucesso
# Retorna uma mensagem de erro em caso de erro
def new_session_esocial(cert, pwd, timeout=20, delay=1):
    url = 'https://login.esocial.gov.br'
    print('\n>>> nova sessão com', cert.split('\\')[-1])

    with pfx_to_pem(cert, pwd) as cert:
        driver = exec_phantomjs(cert=cert, headers=_headers)
        if isinstance(driver, str): return driver

    driver.get(url)
    sleep(1)

    x_path = '/html/body/div[2]/div[3]/form/fieldset/div[2]/div[1]/p/button'
    elem = driver.find_element_by_xpath(x_path)
    elem.click()
    sleep(1)

    for i in range(timeout):
        sleep(delay)
        try:
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            url_main = soup.find('div', attrs={'id':'cert-digital'}).a.get('href', '')
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
_new_session_esocial = new_session_esocial


# Troca de cpf/cnpj no site esocial depois de logado
# Retorna uma string vazia em caso de sucesso
# Retorna uma string mensagem em caso de erro
def new_login_esocial(ni, session):
    print('\n>>> logando', ni)
    url_base = 'https://www.esocial.gov.br/portal'
    url_home = f'{url_base}/Home/IndexProcuracao'
    url_troca = f'{url_base}/Home/Index?trocarPerfil=true'

    ni = ''.join(i for i in ni if i.isdigit())
    data = {
        'perfil': '2', 'trocarPerfil': 'true', 'podeSerMicroPequenaEmpresa': 'False',
        'tipoInscricao': '1', 'EhOrgaoPublico': 'False', 'logadoComCertificadoDigital': 'True',
        'perfilAcesso': '', 'procuradorCpf': '', 'procuradorCnpj': '',
        'representanteCnpj': '', 'inscricao': '', 'inscricaoJudiciario': '',
        'numeroProcessoJudiciario': '',
    }

    if len(ni) == 14:
        f_ni = '{}.{}.{}/{}-{}'.format(ni[:2], ni[2:5], ni[5:8], ni[8:12], ni[12:])
        data['procuradorCnpj'], data['perfilAcesso'] = f_ni, 'PROCURADOR_PJ'
        url_home = f'{url_home}?procuradorCnpj={f_ni}&procuradorCpf='

    elif len(ni) == 11:
        f_ni = '{}.{}.{}-{}'.format(ni[:3], ni[3:6], ni[6:9], ni[9:])
        data['procuradorCpf'], data['perfilAcesso'] = f_ni, 'PROCURADOR_PF'
        url_home = f'{url_home}?procuradorCnpj=&procuradorCpf={f_ni}'

    else:
        return 'Erro Login - cpf/cnpj em formato invalido'

    session.get(url_troca, headers=_headers, verify=False)
    res = session.post(url_home, data, headers=_headers, verify=False)
    
    if res.url == f'{url_base}/Empregador/CadastroDomestico':
        return 'Erro Login - cadastrar dados empregador'

    if res.url == f'{url_base}/Empregador/PrimeiraAdesao':
        return 'Erro Login - primeira adesão'

    if res.url not in _allowed_urls:
        soup = BeautifulSoup(res.content, 'html.parser')
        msg = soup.find('div', attrs={'id':'mensagemGeral'})
        if not msg: return 'Erro Login - verificar login manualmente'

        msg = msg.text.lower().strip().replace('\n', '@').strip("×@@")
        return f'Erro Login - {msg if msg else res.url}'

    return ''
_new_login_esocial = new_login_esocial