# -*- coding: utf-8 -*-
import os, time
from pyautogui import prompt, confirm
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
    
    # abrir o site da JUCESP
    print('>>> Logando na JUCESP')
    url_inicio = 'https://www.jucesponline.sp.gov.br/BuscaAvancada.aspx'
    
    # espera o site abrir
    driver.get(url_inicio)
    while not find_by_id('content', driver):
        time.sleep(1)
    
    # insere a dara de abertura inicial das empresas para pesquisa
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtDataAberturaInicio')
    element.send_keys(data_abertura_inicio)

    # insere a dara de abertura final das empresas para pesquisa
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtDataAberturaFim')
    element.send_keys(data_abertura_final)

    # insere o município das empresas para pesquisa
    element = driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_txtMunicipio')
    element.send_keys(municipio)
    time.sleep(1)
    
    # clica para pesquisar
    driver.find_element(by=By.ID, value='ctl00_cphContent_frmBuscaAvancada_btPesquisar').click()
    
    captcha = _solve_text_captcha(driver, 'ctl00_cphContent_gdvResultadoBusca_pnlCaptcha')
    
    if not captcha:
        print('Erro Login - não encontrou captcha')
    
    # insere a chave do captcha
    driver.find_element(by=By.XPATH, value='//*[@id="formBuscaAvancada"]/table/tbody/tr[1]/td/div/div[2]/label/input').send_keys(captcha)
    time.sleep(1)
    
    # clica para entrar
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
    # pega a "NIRE"
    nire = driver.find_element(by=By.ID, value=id_iten).text
    return nire


def consulta(driver):
    print('>>> Coletando lista de empresas...')
    nire = []
    # entra em um loop infinito e só sai quando não tiver mais páginas com empresas para coletar as "NIREs"
    while 1 > 0:
        # captura as "NIREs" de 2 a 9
        indice = 2
        while find_by_id(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl0{str(indice)}_lbtSelecionar', driver):
            nire_capturada = captura_nire(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl0{str(indice)}_lbtSelecionar', driver)
            nire.append(nire_capturada)
            indice += 1

        # captura as "NIREs" de 10 até acabar
        indice = 10
        while find_by_id(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl{str(indice)}_lbtSelecionar', driver):
            nire_capturada = captura_nire(f'ctl00_cphContent_gdvResultadoBusca_gdvContent_ctl{str(indice)}_lbtSelecionar', driver)
            nire.append(nire_capturada)
            indice += 1
        # tenta mudar de página, se não conseguir a lista acabou e pode sair do loop
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
    # formata as "NIREs"
    nire = str(nire).replace("('", "").replace("',)", "")
    # entra no site colocando a "NIRE" correspondente da empresa para coletar as informações dela
    driver.get(f'https://www.jucesponline.sp.gov.br/Pre_Visualiza.aspx?nire={nire}&idproduto=')
    time.sleep(1)
    
    # espera a página abrir
    while not find_by_id('ctl00_cphContent_frmPreVisualiza_lblEmpresa', driver):
        time.sleep(1)
    
    #
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
    
    # loop para coletar cada iten da empresa, concatena em uma variável e formata para inserir em uma planilha
    infos_da_empresa = ''
    for info in infos:
        info_da_empresa = driver.find_element(by=By.ID, value=str(info[1])).text
        
        infos_da_empresa += info_da_empresa.replace('\n', ' | ').replace(';', ' | ') + ';'

    _escreve_relatorio_csv(f"{nire};{infos_da_empresa}", nome='Consulta JUCESP')
        
    
@_time_execution
def run():
    continuar = confirm(text='Continuar consulta existente?', buttons=('Sim', 'Não'))

    data_abertura_inicio = prompt(text='Data de abertura, de: ', title='Script incrível', default='00/00/0000')
    data_abertura_final = prompt(text='Até: ', title='Script incrível', default='00/00/0000')
    qual_municipio = prompt(text='Município', title='Script incrível', default='')
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    
    # se continuar uma pesquisa pede para abrir a planilha com as "NIREs" e depois a planilha com as empresas já consultadas
    if continuar == 'Sim':
        total_nires = _open_lista_dados()
        index = _where_to_start(tuple(i[0] for i in total_nires))
        if index is None:
            return False
        nires = total_nires[index:]

    # loop para logar no site
    driver = False
    while not driver:
        driver = logar(options, data_abertura_inicio, data_abertura_final, qual_municipio)
        if not driver:
            print('❌ Erro no captcha, tentando novamente...\n')
    
    # função para coletar as "NIREs" no site, só entra aqui se não for continuar uma consulta
    if continuar == 'Não':
        nires = consulta(driver)
    
    for count, nire in enumerate(nires, start=1):
        _indice(count, nires, nire)
        consulta_empresas(nire, driver)
    driver.quit()
        
    
if __name__ == '__main__':
    run()
    print(f'✔ Dados coletados')
    _escreve_header_csv("", nome='Consulta JUCESP')
