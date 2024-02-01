import time, re, os, pyautogui as p
from xhtml2pdf import pisa
from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from captcha_comum import _solve_recaptcha
from fazenda_comum import _atualiza_info
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Fazenda.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.readline()
user = user.split('/')

def login():
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
                     'ctl00$ddlTipoUsuario':  '#rdBtnContabilista',
                     'ctl00$UserNameAcessivel': user[0],
                     'ctl00$PasswordAcessivel': user[1],
                     'ctl00$ConteudoPagina$controleLogin$rblTipo': 'rdBtnContabilista',
                     'ctl00$ConteudoPagina$controleLogin$UserName': user[0],
                     'ctl00$ConteudoPagina$controleLogin$Password': user[1],
                     'g-recaptcha-response': captcha,
                     'ctl00$ConteudoPagina$controleLogin$btnLogin': 'Acessar'}
            
            res = s.post('https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2f', data=query)
            
            soup = BeautifulSoup(res.content, 'html.parser')
            soup = soup.prettify()
            
            if re.compile(r'Falha na resolução do CAPTCHA').search(soup):
                print('>>> Falha na resolução do Captcha, tentando novamente.')
                s.close()
            else:
                if not re.compile(r'Escolha a empresa com a qual deseja operar o sistema').search(soup):
                    return False
                return s


def captura_lista_empresas(driver):
    print('>>> Capturando lista de empresas')
    driver.get('https://www.nfp.fazenda.sp.gov.br/EmissaoNF/cnpjlista.aspx')
    empresas = re.compile(r'href="(javascript:__.+)">\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d').findall(driver.page_source)
    return driver, empresas


def consulta_saldo(driver, empresas, semestre, ano, andamentos):
    print('>>> Consultando saldo das empresas...')
    continua = 'não'
    for count, empresa in enumerate(empresas, start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, empresas)
        if count > 1:
            driver.get('https://www.nfp.fazenda.sp.gov.br/EmissaoNF/cnpjlista.aspx')
            
        driver.execute_script(empresa)
        while not re.compile(r'Bem-vindo ao sistema da Nota Fiscal Paulista').search(driver.page_source):
            if re.compile(r'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto').search(driver.page_source):
                break
            time.sleep(0.1)
        
        if re.compile(r'As funcionalidades do sistema permanecerão indisponíveis enquanto o perfil estiver incompleto').search(driver.page_source):
            cnpj = re.compile(r'ctl00\$ConteudoPagina\$TxtCNPJ\" type=\"text\" value=\"(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)\"').search(driver.page_source).group(1)
            nome = re.compile(r'ctl00\$ConteudoPagina\$TxtRazaoSocial\" type=\"text\" value=\"(.+)\" disabled=\"disabled\"').search(driver.page_source).group(1)
            
            cnpj = re.sub(r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', cnpj)
            nome = nome.replace('&amp;', '&')
            _escreve_relatorio_csv(f'{cnpj};{nome};{semestre};{ano};Cadastro incompleto', nome=andamentos)
            print(f'{cnpj} - {nome}')
            print(f'❌ Cadastro incompleto')
            continue
            
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
            _escreve_relatorio_csv(f'{cnpj};{nome};{semestre};{ano};Período não encontrado', nome=andamentos)
            print(f'❗ Período não encontrado')
            continue
            
        driver.execute_script("document.getElementById('" + id_periodo + "').click()")
        time.sleep(0.2)
        driver.find_element(by=By.ID, value='btnConsultarNFSemestre').click()
        
        while not re.compile(r'Saldo Disponível Para Saque').search(driver.page_source):
            time.sleep(0.1)
        
        saldo = re.compile(r'Saldo Disponível Para Saque:&nbsp;<span id=\"lblSaldo\">(.+)</span></strong>').search(driver.page_source).group(1)
        
        _escreve_relatorio_csv(f'{cnpj};{nome};{semestre};{ano};{saldo}', nome=andamentos)
        print(f'✔ {saldo}')


def create_pdf(driver, nome_arquivo, comp_formatado):
    # salva o pdf criando ele a partir do código html da página, para que o PDF criado seja editável
    e_dir_pdf = os.path.join('execução', 'Arquivos ' + comp_formatado)
    os.makedirs(e_dir_pdf, exist_ok=True)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    # Remove atributo href das tags, links e urls que não serão usados
    for tag in soup.find_all(href=True):
        tag['href'] = ''

    # Remove tags img e script, não serão usados
    _ = list(tag.extract() for tag in soup.find_all('img'))
    _ = list(tag.extract() for tag in soup.find_all('script'))
    
    # remove mais algumas coisas específicas
    soup = str(soup)\
        .replace('target="_blank">Cidadão SP</a></td>', '')\
        .replace('target="_blank">saopaulo.sp.gov.br</a></td>', '') \
        .replace('<a class="botao-sistema" href="">Home</a>', '') \
        .replace('<a class="botao-sistema" href="">Imprimir</a>', '') \
        .replace('<a class="botao-sistema" href="">Encerrar</a>', '')\
        .replace('target="_blank">Ouvidoria</a></td>', '')\
        .replace('target="_blank">Transparência</a></td>', '')\
        .replace('href="" target="_blank">SIC</a></td>', '')
    
    soup = BeautifulSoup(soup, 'html.parser')
    
    with open(os.path.join(e_dir_pdf, nome_arquivo), 'w+b') as pdf:
        pisa.showLogging()
        pisa.CreatePDF(str(soup), pdf)
    
    return driver, 'Arquivo gerado'

    
@_time_execution
def run():
    andamentos = 'Saldo Nota Fiscal Paulista'
    semestre = p.confirm(text='Qual semestre deseja consultar', buttons=['1º Semestre', '2º Semestre'])
    ano = p.prompt(text='Qual ano deseja consultar')
    
    s = login()
    
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
        
        consulta_saldo(driver, empresas, semestre, ano, andamentos)
    else:
        _escreve_relatorio_csv(f'{usuario};{senha};Erro ao logar no usuário')
        
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;NOME;SEMESTRE;ANO;SALDO', nome=andamentos)


if __name__ == '__main__':
    run()
