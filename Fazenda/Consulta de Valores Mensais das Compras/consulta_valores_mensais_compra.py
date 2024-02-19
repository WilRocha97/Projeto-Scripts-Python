import time, re, os, pyautogui as p
from xhtml2pdf import pisa
from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from calendar import monthrange
from datetime import datetime
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from captcha_comum import _solve_recaptcha
from fazenda_comum import _atualiza_info
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _indice, _open_lista_dados, _where_to_start


def login(usuario, senha):
    while True:
        url = 'https://www.nfp.fazenda.sp.gov.br'
        recaptcha_data = {'sitekey': '6LftgKsUAAAAADAPkvLCFfCgnf9OuosqEc19LHnw', 'url': url}
        captcha = _solve_recaptcha(recaptcha_data)
        with Session() as s:
            # entra no site
            pagina = s.get('https://www.nfp.fazenda.sp.gov.br', verify=False)
            
            state, generator, validation = _atualiza_info(pagina)
            
            print('>>> Logando no usuário')
            # loga no usuario
            query = {'__EVENTTARGET': '',
                     '__EVENTARGUMENT': '',
                     '__VIEWSTATE': state,
                     '__VIEWSTATEGENERATOR': generator,
                     '__EVENTVALIDATION': validation,
                     'ctl00$hddIDDoacao': '',
                     'ctl00$ddlTipoUsuario':  '#rdBtnNaoContribuinte',
                     'ctl00$UserNameAcessivel': 'Digite o Usuário',
                     'ctl00$PasswordAcessivel': 'x',
                     'ctl00$ConteudoPagina$controleLogin$rblTipo': 'rdBtnContribuinte',
                     'ctl00$ConteudoPagina$controleLogin$UserName': usuario,
                     'ctl00$ConteudoPagina$controleLogin$Password': senha,
                     'g-recaptcha-response': captcha,
                     'ctl00$ConteudoPagina$controleLogin$btnLogin': 'Acessar'}
            
            res = s.post('https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2f', data=query)
            
            soup = BeautifulSoup(res.content, 'html.parser')
            soup = soup.prettify()
            
            if re.compile(r'Falha na resolução do CAPTCHA').search(soup):
                print('>>> Falha na resolução do Captcha, tentando novamente.')
                s.close()
            else:
                if re.compile(r'Usuário ou senha inválidos').search(soup):
                    print('❌ Usuário ou senha inválidos')
                    s.close()
                    return False, 'Usuário ou senha inválidos'
                
                if re.compile(r'Não existe estabelecimento associado ao usuário ou trata-se de contribuinte produtor rural').search(soup):
                    print(f'❌ Não existe estabelecimento associado ao usuário ou trata-se de contribuinte produtor rural')
                    s.close()
                    return False, 'Não existe estabelecimento associado ao usuário ou trata-se de contribuinte produtor rural'
                
                if re.compile(r'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto').search(soup):
                    print(f'❌ As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto')
                    s.close()
                    return False, 'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto'
                
                if re.compile(r'Escolha a empresa com a qual deseja operar o sistema').search(soup):
                    return s, 'ok'
                
                elif re.compile(r'Bem-vindo ao sistema da Nota Fiscal Paulista').search(soup):
                    return s, 'ok'
                
                return False, 'Erro ao logar no usuário'


def consulta_saldo(driver, periodo, semestre, empresas, cnpj_login, nome_login, andamentos):
    ano = periodo.split('/')[1]
    data_inicial = f'01/{periodo}'
    print(data_inicial)
    # Converta a string para um objeto datetime
    data_obj = datetime.strptime(periodo, "%m/%Y")
    # Obtenha o último dia do mês
    ultimo_dia = monthrange(data_obj.year, data_obj.month)[1]
    # Crie uma nova data usando o último dia
    data_final = data_obj.replace(day=ultimo_dia).strftime("%d/%m/%Y")
    
    print('>>> Consultando saldo das empresas...')
    for count, empresa in enumerate(empresas, start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, empresas)
        if count > 1:
            driver.get('https://www.nfp.fazenda.sp.gov.br/EmissaoNF/cnpjlista.aspx')
            
        if empresa:
            try:
                driver.execute_script(empresa)
            except Exception as erro:
                print(empresa)
                print(erro)
        else:
            driver.get('https://www.nfp.fazenda.sp.gov.br/Principal.aspx')
    
        while not re.compile(r'Bem-vindo ao sistema da Nota Fiscal Paulista').search(driver.page_source):
            if re.compile(r'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto').search(driver.page_source):
                break
            time.sleep(0.1)
        
        if re.compile(r'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto').search(driver.page_source):
            cnpj = re.compile(r'ctl00\$ConteudoPagina\$TxtCNPJ\" type=\"text\" value=\"(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)\"').search(driver.page_source).group(1)
            nome = re.compile(r'ctl00\$ConteudoPagina\$TxtRazaoSocial\" type=\"text\" value=\"(.+)\" disabled=\"disabled\"').search(driver.page_source).group(1)
            
            cnpj = re.sub(r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', cnpj)
            nome = nome.replace('&amp;', '&')
            
            _escreve_relatorio_csv(f'{cnpj_login};{nome_login};{periodo};{cnpj};{nome};Cadastro incompleto', nome=andamentos)
            print(f'{cnpj} - {nome}')
            print(f'❌ Cadastro incompleto')
            continue
        time.sleep(0.1)
        
        # captura dados da empresa
        dados = re.compile(r'CNPJ:</strong>&nbsp;(.+)\n.+\n.+<strong>Razão Social:</strong>&nbsp;(.+)\n.+</p>').search(driver.page_source)
        cnpj = dados.group(1)
        nome = dados.group(2).replace('&amp;', '&')
        
        print(f'{cnpj} - {nome}')
        # abre a tela de consulta
        driver.execute_script(r"javascript:__doPostBack('ctl00$Menu$UCMenu$LoginView1$menuSuperior','Consultar\\CONSULTAR_NF')")
        timer = 0
        while not re.compile(r'Para consultar os documentos fiscais eletrônicos emitidos, acesse:').search(driver.page_source):
            time.sleep(0.1)
            timer += 0.1
            if timer > 20:
                driver.execute_script(r"javascript:__doPostBack('ctl00$Menu$UCMenu$LoginView1$menuSuperior','Consultar\\CONSULTAR_NF')")
        
        print('>>> Abrindo consulta...')
        driver.find_element(by=By.ID, value='rblTipo_1').click()
        
        while not re.compile(r'Semestre de').search(driver.page_source):
            time.sleep(0.1)
        
        print(f'>>> Consultando {semestre} - {ano}')
        id_periodo = re.compile(r'<li><input id=\"(.+)\" type=\"radio\" name=\"ctl00\$ConteudoPagina\$rblPeriodo\" value=.+><label for=.+>' + str(semestre) + ' de ' + str(ano) + '</label></li>').search(driver.page_source).group(1)
        if not id_periodo:
            _escreve_relatorio_csv(f'{cnpj_login};{nome_login};{periodo};{cnpj};{nome};Período não encontrado', nome=andamentos)
            print(f'❗ Período não encontrado')
            continue
            
        driver.execute_script("document.getElementById('" + id_periodo + "').click()")
        time.sleep(0.2)
        driver.find_element(by=By.ID, value='btnConsultaAvancada').click()
        
        while not driver.find_element(by=By.ID, value='txtDataIni'):
            time.sleep(0.1)
        
        driver.execute_script("document.querySelector('#txtDataIni').value='" + str(data_inicial) + "'")
        driver.execute_script("document.querySelector('#txtDataFim').value='" + str(data_final) + "'")
        
        time.sleep(1)
        driver.find_element(by=By.ID, value='btnConsultaNF').click()
        
        saldo = re.compile(r'lblValorTotDoc\">(.+)</span>').search(driver.page_source).group(1)
        
        _escreve_relatorio_csv(f'{cnpj_login};{nome_login};{periodo};{cnpj};{nome};{saldo}', nome=andamentos)
        print(f'✔ {saldo}')


def captura_lista_empresas(driver):
    print('>>> Capturando lista de empresas')
    driver.get('https://www.nfp.fazenda.sp.gov.br/EmissaoNF/cnpjlista.aspx')
    empresas = re.compile(r'href="(javascript:__.+)">\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d').findall(driver.page_source)
    if not empresas:
        empresas = [False]
        
    return driver, empresas


@_time_execution
def run():
    andamentos = 'Valores Mensais de Compra'
    semestre = p.confirm(text='Qual semestre deseja consultar', buttons=['1º Semestre', '2º Semestre'])
    periodo = ''
    while not re.search(r'\d{2}/\d{4}', periodo):
        periodo = p.prompt(text='Competência da consulta:', default='00/0000')
    
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return
        
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, usuario, senha = empresa
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        s, situacao = login(usuario, senha)
        
        if s:
            # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--window-size=1920,1080')
            # options.add_argument("--start-maximized")
        
            status, driver = _initialize_chrome(options)
            driver.get('https://www.nfp.fazenda.sp.gov.br')
            
            for name, value in s.cookies.get_dict().items():
                driver.add_cookie({'name': name, 'value': value})
            
            driver, empresas = captura_lista_empresas(driver)
            
            consulta_saldo(driver, periodo, semestre, empresas, cnpj, nome, andamentos)
        else:
            _escreve_relatorio_csv(f'{cnpj};{nome};{periodo};x;x;{situacao}', nome=andamentos)
        
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;NOME;PERÍODO;CNPJ ENCONTRADO;NOME ENCONTRADO;SALDO', nome=andamentos)


if __name__ == '__main__':
    run()
