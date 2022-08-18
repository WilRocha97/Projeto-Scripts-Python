# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import os, time, re, shutil

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha

os.makedirs('execucao/Certidões', exist_ok=True)


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def find_by_path(xpath, driver):
    try:
        elem = driver.find_element(by=By.XPATH, value=xpath)
        return elem
    except:
        return None


def login(options, cnpj, insc_muni):
    cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
    status, driver = _initialize_chrome(options)
    
    # entra na página inicial da consulta
    url_inicio = 'https://web.jundiai.sp.gov.br/PMJ/SW/certidaonegativamobiliario.aspx'
    driver.get(url_inicio)
    time.sleep(1)
    
    # pega o sitekey para quebrar o captcha
    data = {'url': url_inicio, 'sitekey': '6LfK0igTAAAAAOeUqc7uHQpW4XS3EqxwOUCHaSSi'}
    response = _solve_recaptcha(data)
    
    # insere a inscrição estadual e o cnpj da empresa
    _send_input('DadoContribuinteMobiliario1_txtCfm', insc_muni, driver)
    _send_input('DadoContribuinteMobiliario1_txtNrCicCgc', cnpj, driver)
    
    # pega a id do campo do recaptcha
    id_response = re.compile(r'class=\"recaptcha\".+<textarea id=\"(.+)\" name=')
    id_response = id_response.search(driver.page_source).group(1)
    
    text_response = ''
    while not text_response == response:
        print('>>> Tentando preencher o campo do captcha')
        # insere a solução do captcha via javascript
        driver.execute_script('document.getElementById("' + id_response + '").innerText="' + response + '"')
        time.sleep(2)
        try:
            text_response = re.compile(r'display: none;\">(.+)</textarea></div>')
            text_response = text_response.search(driver.page_source).group(1)
        except:
            pass
    
    # clica em consultar
    print('>>> Carregando consulta')
    driver.find_element(by=By.ID, value='btnEnviar').click()
    time.sleep(1)
    
    while not find_by_id('lblContribuinte', driver):
        if find_by_id('AjaxAlertMod1_lblAjaxAlertMensagem', driver):
            situacao = re.compile(r'AjaxAlertMod1_lblAjaxAlertMensagem\">(.+)</span>')
            situacao = situacao.search(driver.page_source).group(1)
            if situacao == 'Consta(m) pendência(s) para a emissão de certidão por meio da Internet. Dirija-se à Av. União dos Ferroviários, 1760 - Centro - Jundiaí de segunda a sexta-feiras ' \
                           'das 9h:00 às 18h:00 e aos sábados das 9h:00 às 13h:00.':
                situacao_print = f'✔ {situacao}'
                driver.quit()
                return situacao, situacao_print
            if situacao == 'Confirme que você não é um robô':
                situacao_print = '❌ Erro ao logar na empresa'
                driver.quit()
                return 'Erro ao logar na empresa', situacao_print
        time.sleep(1)
    
    print('>>> Consultando empresa')
    situacao = re.compile(r'<span id=\"lblSituacao\">(.+)</span>')
    situacao = situacao.search(driver.page_source).group(1)
    time.sleep(1)
    
    if situacao == 'Não constam débitos para o contribuinte':
        _send_input('txtSolicitante', 'Escritório', driver)
        _send_input('txtMotivo', 'Consulta', driver)
        time.sleep(1)
        driver.find_element(by=By.ID, value='btnImprimir').click()
        
        while not os.path.exists('V:/Setor Robô/Scripts Python/Serviços Prefeitura/Consulta Débitos Municipais Jundiaí/execucao/Certidões/relatorio.pdf'):
            time.sleep(1)
        while os.path.exists('V:/Setor Robô/Scripts Python/Serviços Prefeitura/Consulta Débitos Municipais Jundiaí/execucao/Certidões/relatorio.pdf'):
            try:
                arquivo = f'{cnpj_limpo} - Certidão Negativa de Débitos Municipais Jundiaí.pdf'
                download_folder = "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta Débitos Municipais Jundiaí\\execucao\\Certidões"
                shutil.move(os.path.join(download_folder, 'relatorio.pdf'), os.path.join(download_folder, arquivo))
                time.sleep(2)
            except:
                pass
    
    situacao_print = f'✔ {situacao}'
    driver.quit()
    return situacao, situacao_print


@_time_execution
def run():
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta Débitos Municipais Jundiaí\\execucao\\Certidões",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, insc_muni = empresa

        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        situacao, situacao_print = login(options, cnpj, insc_muni)
        
        _escreve_relatorio_csv(f'{cnpj};{insc_muni};{situacao}', nome='Consulta de débitos municipais de Jundiaí')
        print(situacao_print)


if __name__ == '__main__':
    run()
