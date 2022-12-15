# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import confirm
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome


def localiza_path(driver, elemento):
    try:
        driver.find_element(by=By.XPATH, value=elemento)
        return True
    except:
        return False


def localiza_id(driver, elemento):
    try:
        driver.find_element(by=By.ID, value=elemento)
        return True
    except:
        return False
    

def login_onvio(driver):
    driver.get('https://onvio.com.br/login/#/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # inserir o CNPJ no campo
    driver.find_element(by=By.NAME, value='uid').send_keys('robo@veigaepostal.com.br')
    time.sleep(1)
    
    # inserir a senha no campo
    driver.find_element(by=By.NAME, value='pwd').send_keys('Rb#0086*')
    time.sleep(1)
    
    # clica em acessar
    driver.execute_script('document.querySelector("#trta1-auth > div > button").click()')
    time.sleep(5)
    
    try:
        driver.execute_script('document.querySelector("#trta1-mfs-later").click()')
    except:
        pass
    
    time.sleep(1)
    return driver


def procura_empresa(departamento, empresa, driver):
    numero, email = empresa
    print('>>> Abrindo lista de usuários')
    try:
        driver.get('https://onvio.com.br/br-portal-do-cliente/settings/clients-users')
        
        while not localiza_path(driver, '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-list/div[1]/div[1]/bento-combobox/div[1]/input'):
            time.sleep(1)
        
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-list/div[1]/div[1]/bento-combobox/div[1]/input')\
            .click()
        
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-list/div[1]/div[1]/bento-combobox/div[3]/bento-combobox-list/div[2]/bento-list/cdk-virtual-scroll-viewport/div[1]/div[1]/div[2]')\
            .click()
        
        time.sleep(1)
        # clica para abrir a barra de pesquisa
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-list/div[2]/on-toolbar/bento-toolbar/div[1]/ul/li/input')\
            .send_keys(email)
    except:
        driver.close()
        login_onvio(driver)
        return 'erro', 'fechado'
    
    time.sleep(2)

    # Sua busca não retornou nenhum resultado'
    try:
        nao_encontrada = re.compile(r'__title\"> (Sua busca não retornou nenhum resultado)\. </h2>').search(driver.page_source).group(1)
    except:
        nao_encontrada = ''
        
    if nao_encontrada == 'Sua busca não retornou nenhum resultado':
        _escreve_relatorio_csv(f'{numero};{email};E-mail não encontrado')
        print('❗ E-mail não encontrado')
        return 'continue', driver
    
    try:
        # clica para selecionar a empresa
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-list/div[2]/on-grid/div[1]/div[1]/div[1]/div[2]/div[1]/div/span[1]')\
            .click()
        
        time.sleep(1)
        while not localiza_path(driver, '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/button'):
            time.sleep(1)
    
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/button') \
            .click()
    except:
        driver.close()
        login_onvio(driver)
        return 'erro', 'fechado'
    
    time.sleep(1)
    print('>>> Editando departamentos')

    try:
        unchecked = re.compile(r'(aria-checked=\"true\")><i aria-hidden=\"true\" class=\"bui-checkbox-unchecked\"></i><i aria-hidden=\"true\" class=\"bui-checkbox-checked\"></i></bento-checkbox><div class=\"bento-multiselect-list-item-label\"> Selecionar todos') \
            .search(driver.page_source).group(1)
    except:
        unchecked = 'unchecked'

    if unchecked == 'unchecked':
        # clica em Selecionar todos
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[1]') \
        .click()
    
    # clique para tirar a seleção de todos
    time.sleep(0.2)
    driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[1]') \
    .click()

    time.sleep(1)
    if departamento == 'Geral':
        botoes = [('Administrativo', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[3]'),
                       ('Cliente contábil', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[4]'),
                       ('Contábil', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[6]'),
                       ('Financeiro', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[8]'),
                       ('Fiscal', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[9]'),
                       ]
    
    elif departamento == 'Departamento pessoal':
        botoes = [('Cliente folha', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[5]', 'CLIENTE FOLHA'),
                       ('Departamento pessoal', '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/bento-multiselect-list/bento-list/cdk-virtual-scroll-viewport/div[1]/div[7]'),
                       ]
    
    else:
        return 'continue', driver

    _escreve_relatorio_csv(f'{numero};{email}', end=';')
    for botao in botoes:
        
        if departamento == 'Departamento pessoal':
            try:
                re.compile(r'title=\"' + str(botao[2]) + '\"><bento-checkbox class=\"ng-untouched ng-pristine ng-valid\"><input type=\"checkbox\" class=\"ng-untouched ng-pristine ng-valid\" tabindex=\"-1\" (aria-checked=\"true\")>') \
                .search(driver.page_source).group(1)
                continue
            except:
                pass
        
        # seleciona o botão
        driver.find_element(by=By.XPATH, value=str(botao[1])).click()
        _escreve_relatorio_csv('Ok', end=';')
        print(f'✔ {botao[0]}')

    _escreve_relatorio_csv('')
    # clica em concluir
    driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/div/div[1]/app-general-data-client/div/form/div[3]/div[1]/app-departments/div/div[1]/bento-multiselect-overlay/div/div/div/div[2]/div/button[1]') \
    .click()

    print('>>> Concluíndo alteração')
    time.sleep(2)
    # clica em avançar
    i = 1
    while i <= 5:
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-client-user-container/div/app-client-user-edit/div[1]/on-wizard-add/footer/button[2]') \
        .click()
        time.sleep(0.2)
        i += 1

    time.sleep(2)
    print('✔ Alteração concluída')
    return 'continue', driver
    

@_time_execution
def run():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    departamento = confirm(title='Script incrível!', text='Qual departamento configurar?', buttons=('Geral', 'Departamento pessoal'))
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    
    # fazer login no Onvio
    login_onvio(driver)

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        
        erro = 'erro'
        while erro == 'erro':
            erro, driver = procura_empresa(departamento, empresa, driver)
        
    driver.close()
    _escreve_header_csv(texto='E-MAIL;ADMINISTRATIVO;CLIENTE CONTÁBIL;CONTABIL;FINANCEIRO;FISCAL')
    
if __name__ == '__main__':
    run()
