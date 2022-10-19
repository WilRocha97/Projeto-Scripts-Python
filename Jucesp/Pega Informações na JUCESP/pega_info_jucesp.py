# -*- coding: utf-8 -*-
import os, time
from pyautogui import prompt, confirm
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _indice, _open_lista_dados, _where_to_start
from captcha_comum import _solve_text_captcha

os.makedirs('execucao', exist_ok=True)


def find_by_id(id_iten, driver):
    try:
        elem = driver.find_element(by=By.ID, value=id_iten)
        return elem
    except:
        return None


def find_by_path(xpath, driver):
    try:
        elem = driver.find_element(by=By.XPATH, value=xpath)
        return elem
    except:
        return None


def logar(options, data_abertura_inicio, data_abertura_final, municipio):
    print(data_abertura_inicio, data_abertura_final, municipio)
    status, driver = _initialize_chrome(options)
    
    print('>>> Logando na JUCESP')
    url_inicio = 'https://www.jucesponline.sp.gov.br/BuscaAvancada.aspx'
    
    driver.get(url_inicio)
    while not find_by_id('content', driver):
        time.sleep(1)
    
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtDataAberturaInicio')
    element.send_keys(data_abertura_inicio)
    
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtDataAberturaFim')
    element.send_keys(data_abertura_final)
    
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtMunicipio')
    element.send_keys(municipio)
    time.sleep(1)
    
    driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_btPesquisar').click()
    
    while not find_by_id('ctl00_cphContent_gdvResultadoBusca_pnlCaptcha', driver):
        time.sleep(1)
    element = driver.find_element(by=By.XPATH, value='//*[@id="formBuscaAvancada"]/table/tbody/tr[1]/td/div/div[1]/img')
    location = element.location
    size = element.size
    driver.save_screenshot('ignore\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    time.sleep(2)
    im = Image.open(r'ignore\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'ignore\captcha\captcha.png')
    time.sleep(1)
    captcha = _solve_text_captcha(os.path.join('ignore', 'captcha', 'captcha.png'))
    
    if not captcha:
        print('Erro Login - não encontrou captcha')
    
    driver.find_element(by=By.XPATH, value='//*[@id="formBuscaAvancada"]/table/tbody/tr[1]/td/div/div[2]/label/input').send_keys(captcha)
    time.sleep(1)
    
    driver.find_element(by=By.ID, value='ctl00_cphContent_gdvResultadoBusca_btEntrar').click()
    
    timer = 0
    print('>>> Aguarde...')
    while not find_by_id('ctl00_cphContent_gdvResultadoBusca_gdvContent', driver):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False
        
    return driver


def captura_nire(id_iten, driver):
    nire = driver.find_element(by=By.ID, value=id_iten).text
    return nire


def consulta(driver):
    print('>>> Coletando lista de empresas...')
    nire = []
    while 1 > 0:
        indice = 2
        while find_by_id(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl0{str(indice)}_lbtSelecionar', driver):
            nire_capturada = captura_nire(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl0{str(indice)}_lbtSelecionar', driver)
            nire.append(nire_capturada)
            indice += 1
        
        indice = 10
        while find_by_id(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl{str(indice)}_lbtSelecionar', driver):
            nire_capturada = captura_nire(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl{str(indice)}_lbtSelecionar', driver)
            nire.append(nire_capturada)
            indice += 1
        try:
            driver.find_element(by=By.ID, value='ctl00_cphContent_gdvResultadoBusca_pgrGridView_btrNext_lbtText').click()
            time.sleep(1)
        except:
            break

    print('❕ Lista coletada\n')
    _escreve_relatorio_csv(str(nire).replace('[', '').replace(']', '').replace("'", "").replace(', ', '\n'), nome='NIRES JUCESP')
    return nire
    

def consulta_empresas(nire, driver):
    print('>>> Coletando dados das empresas')
    nire = str(nire).replace("('", "").replace("',)", "")
    driver.get(f'https://www.jucesponline.sp.gov.br/Pre_Visualiza.aspx?nire={nire}&idproduto=')
    time.sleep(1)
    
    while not find_by_id('ctl00_cphContent_frmPreVisualiza_lblEmpresa', driver):
        time.sleep(1)
    
    infos = [('cnpj', 'ctl00_cphContent_frmPreVisualiza_lblCnpj'),
             ('nome', 'ctl00_cphContent_frmPreVisualiza_lblEmpresa'),
             ('inscricao', 'ctl00_cphContent_frmPreVisualiza_lblInscricao'),
             ('tipo', 'ctl00_cphContent_frmPreVisualiza_lblDetalhes'),
             ('data_constituicao', 'ctl00_cphContent_frmPreVisualiza_lblConstituicao'),
             ('data_ativiade', 'ctl00_cphContent_frmPreVisualiza_lblAtividade'),
             ('objeto', 'ctl00_cphContent_frmPreVisualiza_lblObjeto'),
             ('capital', 'ctl00_cphContent_frmPreVisualiza_lblCapital'),
             ('logradouro', 'ctl00_cphContent_frmPreVisualiza_lblLogradouro'),
             ('numero', 'ctl00_cphContent_frmPreVisualiza_lblNumero'),
             ('complemento', 'ctl00_cphContent_frmPreVisualiza_lblComplemento'),
             ('bairro', 'ctl00_cphContent_frmPreVisualiza_lblBairro'),
             ('cep', 'ctl00_cphContent_frmPreVisualiza_lblCep'),
             ('municipio', 'ctl00_cphContent_frmPreVisualiza_lblMunicipio'),
             ('uf', 'ctl00_cphContent_frmPreVisualiza_lblUf')]

    infos_da_empresa = ''
    for info in infos:
        info_da_empresa = driver.find_element(by=By.ID, value=str(info[1])).text
        
        infos_da_empresa += info_da_empresa.replace('\n', ' | ').replace(';', ' | ') + ';'

    _escreve_relatorio_csv(f"{nire};{infos_da_empresa}", nome='Consulta JUCESP')
        
    
@_time_execution
def run():
    continuar = confirm(text='Continuar consulta existente?', buttons=('Sim', 'Não'))
    data_abertura_inicio = '01/01/2022'
    data_abertura_final = '15/01/2022'
    qual_municipio = 'valinhos'
    nires = ''
    if continuar == 'Não':
        data_abertura_inicio = prompt(text='Data de abertura, de: ', title='Script incrível', default='00/00/0000')
        data_abertura_final = prompt(text='Até: ', title='Script incrível', default='00/00/0000')
        qual_municipio = prompt(text='Município', title='Script incrível', default='')
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    
    if continuar == 'Sim':
        total_nires = _open_lista_dados()
        index = _where_to_start(tuple(i[0] for i in total_nires))
        if index is None:
            return False
        nires = total_nires[index:]

    driver = False
    while not driver:
        driver = logar(options, data_abertura_inicio, data_abertura_final, qual_municipio)
        if not driver:
            print('❌ Erro no captcha, tentando novamente...\n')

    if continuar == 'Não':
        nires = consulta(driver)
                
    for count, nire in enumerate(nires, start=1):
        _indice(count, nires, nire)
        consulta_empresas(nire, driver)
    driver.quit()
        
    print(f'✔ Dados coletados')
    _escreve_header_csv("", nome='Consulta JUCESP')
    

if __name__ == '__main__':
    run()
