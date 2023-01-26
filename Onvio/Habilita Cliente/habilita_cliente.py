# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import confirm
from bs4 import BeautifulSoup
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers, _escreve_header_csv
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
    

def verifica_botao(driver, caminho):
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    html = soup.prettify()
    try:
        re.compile(caminho).search(html).group(1)
        return True
    except:
        return False


def login_onvio(driver):
    driver.get('https://onvio.com.br/login/#/')
    print('>>> Acessando o site...')
    time.sleep(1)

    dados = "V:\\Setor Robô\\Scripts Python\\Onvio\\Dados.txt"
    f = open(dados, 'r', encoding='utf-8')
    user = f.read()
    user = user.split('/')
    
    # inserir o CNPJ no campo
    driver.find_element(by=By.NAME, value='uid').send_keys(user[0])
    
    # inserir a senha no campo
    driver.find_element(by=By.NAME, value='pwd').send_keys(user[1])
    time.sleep(1)
    
    # clica em acessar
    driver.execute_script('document.querySelector("#trta1-auth > div > button").click()')
    time.sleep(5)
    
    try:
        driver.execute_script('document.querySelector("#trta1-mfs-later").click()')
    except:
        pass

    while not localiza_path(driver, '/html/body/bm-optional-header/bm-staff-nav/bm-nav/nav/ul/li[1]/a'):
        time.sleep(1)
        
    print('>>> Abrindo lista de usuários...\n')
    driver.get('https://onvio.com.br/br-portal-do-cliente/cnd/enabling-clients')
    
    time.sleep(1)
    return driver


def procura_empresa(empresa, driver):
    numero, cnpj = empresa
    try:
        while not localiza_path(driver, '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-client-enabling/div/div[1]/on-toolbar/bento-toolbar/div[2]/ul/li/input'):
            time.sleep(1)
        
        # insere o cnpj
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-client-enabling/div/div[1]/on-toolbar/bento-toolbar/div[2]/ul/li/input')\
            .send_keys(cnpj)
    except:
        return 'erro' , driver
    
    time.sleep(2)
    # Sua busca não retornou nenhum resultado'
    try:
        nao_encontrada = re.compile(r'__title\"> (Sua busca não retornou nenhum resultado)\. </h2>').search(driver.page_source).group(1)
    except:
        nao_encontrada = ''
        
    if nao_encontrada == 'Sua busca não retornou nenhum resultado':
        _escreve_relatorio_csv(f'{numero};{cnpj};CNPJ não encontrado')
        print('❗ CNPJ não encontrado')
        driver.get('https://onvio.com.br/br-portal-do-cliente/cnd/enabling-clients')
        return 'continue', driver
    
    # clicar para entrar no cliente
    driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-client-enabling/div/div[1]/on-grid/div[1]/div[1]/div[1]/div[2]/div[2]/div/span') \
        .click()
    
    # espera a página carregar
    print('>>> Abrindo perfil do cliente...')
    while not localiza_path(driver, '/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-issuing-parameters/form/div[1]/div[2]/button[1]'):
        time.sleep(1)
    
    situacao, driver = habilita_agencias(empresa, driver)
    return situacao, driver


def habilita_agencias(empresa, driver):
    numero, cnpj = empresa
    time.sleep(1)
    print('>>> Habilitando CNPJ...')
    
    html = BeautifulSoup(driver.page_source, 'html.parser').prettify()
    # print(html)
    indices = coleta_agencia_indices(html)
    
    try:
        botoes_agencia = re.compile(r'\"name-agency\">\n\s+(.+)\n.+\n.+\n.+(aria-checked=\"false\")').findall(html)
        for botao in botoes_agencia:
            nome_botao = botao[0]
            timer = 1
            while verifica_botao(driver, r'\"name-agency\">\n\s+(' + nome_botao + ')\n.+\n.+\n.+(aria-checked=\"false\")'):
                # item
                driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-issuing-parameters/form/div[2]/div[2]/div[' + str(indices[nome_botao]) + ']/div[2]/bento-toggle').click()
                time.sleep(0.1)
                timer += 1
                if timer >= 30:
                    return 'erro', driver
        
        time.sleep(1)
        html = BeautifulSoup(driver.page_source, 'html.parser').prettify()
        
        botoes_disponibilizar = re.compile(r'\"name-agency\">\n\s+(.+)(\n.+){35}<bento-checkbox.+\n.+(input aria-checked=\"false\")').findall(html)
        for botao in botoes_disponibilizar:
            nome_botao = botao[0]
            timer = 1
            while verifica_botao(driver, r'\"name-agency\">\n\s+(' + nome_botao + ')(\n.+){35}<bento-checkbox.+\n.+(input aria-checked=\"false\")'):
                # disponibilizar para o cliente
                driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-issuing-parameters/form/div[2]/div[2]/div[' + str(indices[nome_botao]) + ']/div[3]/div[2]/div/label/bento-checkbox/input').click()
                time.sleep(0.1)
                if timer >= 30:
                    return 'erro', driver
                
        # salvar
        time.sleep(1)
        driver.find_element(by=By.XPATH, value='/html/body/app-root/div/div/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/main/app-cnd-container/div/app-cnd-issuing-parameters/form/div[1]/div[2]/button[1]') \
            .click()
    
        time.sleep(2)
        _escreve_relatorio_csv(f'{numero};{cnpj};CNPJ habiliado')
        print('✔ CNPJ habiliado')
        return 'continue', driver
    
    except:
        return 'erro', driver
    
    
def coleta_agencia_indices(html):
    agencias = re.compile(r'\"name-agency\">\n\s+(.+)').findall(html)
    
    indice = 1
    indices = {}
    for agencia in agencias:
        indices[agencia]= indice
        indice += 1
    return indices
    
    
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
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,3080')
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
    driver = login_onvio(driver)

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        numero, cnpj = empresa
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        
        erro = 'erro'
        tentativas = 1
        while erro == 'erro':
            
            erro, driver = procura_empresa(empresa, driver)
            if erro == 'erro':
                print('❗ Erro na habilitação da empresa, tentando novamente...\n')
                driver.close()
                # iniciar o driver do chome
                status, driver = _initialize_chrome(options)
    
                # fazer login no Onvio
                driver = login_onvio(driver)
            tentativas += 1
            
            if tentativas >= 3:
                _escreve_relatorio_csv(f'{numero};{cnpj};Erro')
                print('❌ Erro')
                break
            
    driver.close()


if __name__ == '__main__':
    run()
