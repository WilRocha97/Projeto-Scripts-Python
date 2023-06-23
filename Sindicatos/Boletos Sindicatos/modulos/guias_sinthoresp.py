# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re, pyperclip, pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _escreve_relatorio_csv
from pyautogui_comum import _find_img, _click_img, _wait_img


def login(driver, cnpj, senha):
    # entra no site
    driver.get('https://app.sinthoresp.org.br/boletos/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # espera o campo do cnpj aparecer
    while not _find_by_id('CnpjEmpresa', driver):
        time.sleep(1)
    time.sleep(1)
    
    # insere o cnpj da empresa
    driver.find_element(by=By.ID, value='CnpjEmpresa').send_keys(cnpj)
    
    # insere a senha da empresa
    driver.find_element(by=By.ID, value='SenhaEmpresa').send_keys(senha)
    
    # clica em continuar
    driver.find_element(by=By.ID, value='btacessa').click()
    
    while not _find_by_path('/html/body/div[4]/h2', driver):
        try:
            # mude para o contexto do alerta
            alert = driver.switch_to.alert
            # confirme o alerta
            alert.accept()
        except:
            pass
        
    print('>>> Verificando boletos')
    
    avisos = re.compile(r'ATENÇÃO! (.+).<br> ').search(driver.page_source)
    if avisos:
        avisos = avisos.group(1)
        avisos = avisos.replace('<br>', '')
        return driver, avisos
    
    return driver, ''

def gera_boleto(driver, avisos, cnpj, valor, valor_rem, valor_rec, data, funcionarios, responsavel, email):
    data = data.split('-')
    data = f'{data[1]}/{data[0]}'
    # clica em contribuição assistencial
    driver.find_element(by=By.ID, value='contrib2').click()
    
    # espera aparecer a tabela
    while not _find_by_path('/html/body/div[4]/b[2]/font/div[2]/table', driver):
        time.sleep(1)
    
    # verifica se existe um boleto com a data especificada na planilha
    data = re.compile(r'>(' + data + ')<').search(driver.page_source)
    if data:
        data = data.group(1)
        # captura o nome do atributo onclick referente a data especificada
        onclick = re.compile(r'>' + data + '<.+onclick=\"(.+;)\"').search(driver.page_source).group(1)
        # clica no botão para emitir o boleto
        driver.find_element(by=By.XPATH, value="//a[@onclick='" + onclick + "']").click()
        
        # aguarda a tela para inserir o valor do boleto
        while not _find_by_id('VlDocto', driver):
            time.sleep(1)
            
            if _find_by_id('EMPAC778__ENTREGUE_POR', driver):
                lista_campos = [('EMPAC778__ENTREGUE_POR', responsavel), ('EMPAC778__EMAIL', email), ('EMPAC778__QTD_EMPREG', funcionarios),
                                ('EMPAC778__VL_REMUNERACAO', valor_rem), ('EMPAC778__VL_RECOLHIDO', valor_rec)]
                
                for campo in lista_campos:
                    # insere o responsável, email, quantidade de funcionarios, valor de remuneração, valor recolhido
                    driver.find_element(by=By.ID, value=campo[0]).send_keys(campo[1])
        
        
        return driver, f'em construção - {avisos}'
        # return driver, f'✔ Boleto gerado - {avisos}'
    else:
        return driver, f'❗ Boleto não disponível - {avisos}'


def run(empresa):
    # define as variáveis que serão usadas
    cnpj = empresa[1]
    valor_boleto = empresa[2]
    valor_rem = empresa [3]
    valor_rec = empresa[4]
    data = empresa[5]
    senha = empresa[7]
    funcionarios = empresa[8]
    responsavel = empresa[9]
    email = empresa[10]
    
    print('58 - SINTHORESP - Sindicato dos Trabalhadores em Hotéis, Apart Hotéis, Motéis, Flats, Restaurantes, Bares, Lanchonetes e Similares de São Paulo e Região')
    
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    #options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")

    status, driver = _initialize_chrome(options)
    
    # faz login na empresa
    driver, avisos = login(driver, cnpj, senha)
    if not driver:
        driver.close()
        return f'❌ Erro no login - {avisos}'
        
    # gera os boletos
    driver, avisos = gera_boleto(driver, avisos, cnpj, valor_boleto, valor_rem, valor_rec, data, funcionarios, responsavel, email)
    driver.close()
    return f'✔ Boleto gerado - {avisos}'
