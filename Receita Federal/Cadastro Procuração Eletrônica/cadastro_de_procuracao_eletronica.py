# -*- coding: utf-8 -*-
from selenium.common.exceptions import WebDriverException
from pyautogui import alert, prompt
from selenium import webdriver
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from datetime import date
from time import sleep
from sys import path
from PIL import Image
import os, sys

if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    print('>>> running in a PyInstaller bundle')
    sys.path.append(sys._MEIPASS)
    bundle_src = sys._MEIPASS
else:
    print('>>> running in a normal Python process')
    bundle_src = os.path.abspath(os.path.dirname(__file__))
    sys.path.append(os.path.abspath(os.path.join(bundle_src, os.pardir)))

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome, _send_select, _send_input


# obs.: transpor a sessão do webdriver para o request não funciona
# por conta do captcha

_title = 'consulta proc. eletrônica'


def find_by_id(id_do_item, driver):
    try:
        elem = driver.find_element(by=By.ID, value=id_do_item)
        return elem
    except:
        return None


def salva_proc_pdf(id_proc, html):
    os.makedirs('execucao\docs', exist_ok=True)
    nome_arq = os.path.join('execucao\docs', f'proc_eletrônica_{id_proc}.pdf')

    soup = BeautifulSoup(html, 'html.parser')
    div = soup.find('div', attrs={'id': 'rfb-main-container'})
    div.find('fieldset', attrs={'class': 'fieldset-fill'}).extract()
    text_html = "<html><head></head><body>" + str(div) + "</body></html>"

    with open(nome_arq, 'w+b') as pdf:
        pisa.showLogging()
        pisa.CreatePDF(text_html, pdf)


def consulta_proc(tipo, dados, driver):
    url_home = 'https://servicos.receita.fazenda.gov.br/Servicos/procuracoesrfb/controlador/controlePrincipal.asp?acao=telaInicial'
    sitekey = '6LcM8oMhAAAAABaLv_NpRn2EEcBTNQp_WDlpYDDJ'
    driver.get(url_home)

    while not find_by_id('g-recaptcha-response', driver):
        sleep(1)
        
    recaptcha_data = {'sitekey': sitekey, 'url': url_home}
    response = _solve_recaptcha(recaptcha_data)

    # insere a solução do captcha via javascript
    driver.execute_script('document.getElementById("g-recaptcha-response").innerText="' + response + '"')
    sleep(2)
    
    driver.find_element(by=By.ID, value='btnContinuar').click()

    print('>>> Aguardando o site responder')
    while True:
        sleep(1)
        source = driver.page_source
        soup = BeautifulSoup(source, 'html.parser')

        elem = soup.find('span', attrs={'class': 'aviso'})
        if elem:
            text = elem.text.strip()
            if text != '':
                if 'Código incorreto' in text:
                    continue
                return text.split('.')[0]

        elem = soup.find('input', attrs={'id': 'tipoDelegante'})
        if elem:
            break

    print('>>> Preenchendo formulário outorgante')
    if tipo == 'pf':
        id_nasc_cpf = 'nacionalidadeDelegante'
        x_path = '/html/body/div[2]/div[2]/div[1]/div/div/div/form/fieldset/fieldset[1]/label[1]/input'
    else:
        id_nasc_cpf = 'cpfRespLegalDelegante'
        x_path = '/html/body/div[2]/div[2]/div[1]/div/div/div/form/fieldset/fieldset[1]/label[2]/input'

    elem = driver.find_element(by=By.XPATH, value=x_path)
    if elem:
        elem.click()

    logra = ''.join(i for i in dados[1] if i.isalnum() or i in (' ', ','))
    _send_input('delegID', dados[0], driver)
    _send_input('delegEnderecoLogradouro', logra, driver)
    _send_input('delegEnderecoCidade', dados[2], driver)
    _send_input('delegEnderecoEstado', dados[3], driver)
    _send_input('delegTelefone', dados[4], driver)
    _send_input('delegRg', dados[5], driver)
    _send_input(id_nasc_cpf, dados[6], driver)
    _send_input('delegOrgaoExpedidor', dados[7], driver)

    print('>>> Preenchendo formulário outorgado')
    x_path = '/html/body/div[2]/div[2]/div[1]/div/div/div/form/fieldset/fieldset[2]/label[2]/input'
    elem = driver.find_element(by=By.XPATH, value=x_path)
    if elem:
        elem.click()

    # na teoria essas informações não mudam, RPEM
    _send_input('procID', '35586086000160', driver)
    _send_input('procEnderecoLogradouro', 'Rua Fioravante Basílio Maglio, nº 258, Sala 5, Vila Nova Valinhos', driver)
    _send_input('procEnderecoCidade', 'Valinhos', driver)
    _send_input('procEnderecoEstado', 'SP', driver)
    _send_input('procTelefone', '1938716211', driver)
    _send_input('procRg', '250309075', driver)
    _send_input('cpfRespLegalProcurador', '27329597821', driver)
    _send_input('procOrgaoExpedidor', 'SSP.SP.', driver)

    _send_input('codigoAcesso', 'Veiga123', driver)
    _send_input('codigoAcessoRepeticao', 'Veiga123', driver)

    data_i = date.today()
    data_f = date(data_i.year + 5, data_i.month, data_i.day - 1)
    _send_input('dataVigenciaInicio', data_i.strftime('%d/%m/%Y'), driver)
    _send_input('dataVigenciaFim', data_f.strftime('%d/%m/%Y'), driver)

    x_path = '/html/body/div[2]/div[2]/div[1]/div/div/div/form/fieldset/fieldset[4]/label/input'
    elem = driver.find_element(by=By.XPATH, value=x_path)
    if elem:
        elem.click()

    elem = driver.find_element(by=By.ID, value='BotaoCadastrar')
    if elem:
        elem.click()

    while True:
        sleep(1)
        try:
            elem = driver.find_element(by=By.ID, value='divMensagem')
            text = elem.get_attribute('innerText').strip()
        except:
            if "SOLICITAÇÃO DE PROCURAÇÃO PARA A SECRETARIA" not in driver.page_source:
                text = 'Erro ao solicitar a procuração'
                continue
            try:
                salva_proc_pdf(dados[0], driver.page_source)
                text = 'Procuração cadastrada'
            except Exception as e:
                text = 'Erro ao salvar procuração em pdf'
                print(e)
        break

    return text


def valida_dados(dados):
    lista_uf = (
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO',
        'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI',
        'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
    )

    text = []
    if len(dados[5]) in (9, 10):
        text.append('rg')

    if len(dados[0]) in (11, 14):
        text.append('cnpj')

    if len(dados[4]) in (10, 11):
        text.append('telefone')

    if dados[3].upper() in lista_uf:
        text.append('sigla estado')

    return (True, '') if text else (False, "Dados inválidos, verifique:" + ", ".join(text))


@_time_execution
def run():
    status, driver = _initialize_chrome()
    if not status:
        alert(title=_title, text=driver)
        return False

    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        _indice(count, total_empresas, empresa, index)

        status, situacao = valida_dados(empresa)
        if status:
            tipo = 'pj' if len(empresa[0]) == 14 else 'pf'
            '''try:'''
            situacao = consulta_proc(tipo, empresa, driver)
            '''except (BaseException, WebDriverException):
                texto = "Chrome foi encerrado manualmente ou por um processo indevido"
                alert(title=_title, text=texto)
                break'''

        texto = f'{empresa[0]};{situacao}'
        _escreve_relatorio_csv(texto)
        print(f'>>> {texto}')

    driver.quit()

    text = f"Processo concluído com sucesso os arquivos estão em {os.getcwd()}"
    alert(title=_title, text=text)
    return True


if __name__ == '__main__':
    run()
