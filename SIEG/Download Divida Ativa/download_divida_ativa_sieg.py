# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import prompt, confirm
import os, time, re, csv, shutil, fitz, pdfkit

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
    

def login_sieg(driver):
    driver.get('https://auth.sieg.com/')
    print('>>> Acessando o site')
    time.sleep(1)

    dados = "V:\\Setor Robô\\Scripts Python\\Veri\\Dados.txt"
    f = open(dados, 'r', encoding='utf-8')
    user = f.read()
    user = user.split('/')
    
    # inserir o emailno campo
    driver.find_element(by=By.ID, value='txtEmail').send_keys(user[0])
    time.sleep(1)
    
    # inserir a senha no campo
    driver.find_element(by=By.ID, value='txtPassword').send_keys(user[1])
    time.sleep(1)
    
    # clica em acessar
    driver.find_element(by=By.ID, value='btnSubmit').click()
    time.sleep(1)
    
    return driver
    

def sieg_iris(driver):
    print('>>> Acessando IriS DCTF WEB')
    driver.get('https://hub.sieg.com/IriS/#/DividaAtiva')
    return driver


def consulta_lista(driver):

    print('>>> Consultando lista de arquivos')
    
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span'):
        driver.execute_script('document.getElementById("modalYouTube").style="display: none;"')
        time.sleep(1)

    driver = sieg_iris(driver)

    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span'):
        time.sleep(1)
        
    paginas = re.compile(r'>(\d+)</a></span><a class=\"paginate_button btn btn-default next').search(driver.page_source).group(1)
    
    for pagina in range(int(paginas)+1):
        if pagina == 0:
            continue
        
        while not espera_pagina(pagina, driver):
            time.sleep(1)
        
        print(f'>>> Pagina: {pagina}')
        
        detalhes = re.compile(r'btn-details float-right\" (data-id=\".+\") onclick=\"SeeDetails').findall(driver.page_source)
        for detalhe in detalhes:
            driver.find_element(by=By.XPATH, value=f'//a[@{detalhe}]').click()
        
        time.sleep(1)
        # pega a lista de guias da competência desejada
        lista_divida = re.compile(r'excel\"><a id=\"(.+)\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"').findall(driver.page_source)
    
        contador = 0
        erro = ''
        sem_recibo = ''
        # faz o download dos comprovantes
        for divida in lista_divida:
            print('>>> Tentando baixar o arquivo')
            download = True
            while download:
                driver, contador, erro, download= download_divida(contador, driver, divida)
        
        # _escreve_relatorio_csv(f'{cnpj};{nome};Comprovantes {competencia} baixados;{contador} Arquivos;{sem_recibo}')
        
        if pagina != int(paginas):
            driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
        
        time.sleep(1)
        
    return driver


def espera_pagina(pagina, driver):
    try:
        re.compile(r'paginate_button btn btn-default current\" aria-controls=\"tableActiveDebit\" data-dt-idx=\"' + str(pagina) + '\" tabindex').search(driver.page_source)
        return True
    except:
        return False
    

def download_divida(contador, driver, divida):
    while not click(driver,divida):
        time.sleep(5)
    
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa"
    final_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\execução\\Divida Ativa"
    
    contador_3 = 0
    while os.listdir(download_folder) == []:
        if contador_3 > 10:
            while not click(driver, divida):
                time.sleep(5)
            contador_3 = 0
        time.sleep(1)
        contador_3 += 1
    
    for arquivo in os.listdir(download_folder):
        # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
        if re.compile(r'.tmp').search(arquivo):
            os.remove(os.path.join(download_folder, arquivo))
            return driver, contador, 'ok', True
        
        while re.compile(r'crdownload').search(arquivo):
            print('>>> Aguardando download...')
            time.sleep(3)
            for arq in os.listdir(download_folder):
                arquivo = arq
            
        else:
            converte_html_pdf(download_folder, final_folder, arquivo)
            time.sleep(2)
            print(f'✔ {arquivo}')
            contador += 1
            
            for arquivo in os.listdir(download_folder):
                os.remove(os.path.join(download_folder, arquivo))
                
    return driver, contador, 'ok', False


def click(driver, divida):
    print('>>> Baixando arquivo')
    contador_2 = 0
    clicou = 'não'
    while clicou == 'não':
        '''try:'''
        driver.find_element(by=By.ID, value=divida).click()
        clicou = 'sim'
        '''except:
            contador_2 += 1
            clicou = 'não'''
        if contador_2 > 5:
            return False
    
    return True


def converte_html_pdf(download_folder, final_folder, arquivo):
    novo_arquivo = ''
    
    # Defina o caminho para o arquivo HTML e PDF
    caminho_html = os.path.join(download_folder, arquivo)
    caminho_pdf = os.path.join(final_folder, novo_arquivo)
    
    # Converta o arquivo HTML para PDF
    pdfkit.from_file(caminho_html, caminho_pdf)


@_time_execution
def run():
    os.makedirs('execução/Divida Ativa', exist_ok=True)
    os.makedirs('ignore/Divida Ativa', exist_ok=True)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": False  # It will not show PDF directly in chrome
    })
    
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)

    erro = 'sim'
    while erro == 'sim':
        '''try:'''
        driver = sieg_iris(driver)
        driver = consulta_lista(driver)
        erro = 'não'
        '''except:
            erro = 'sim'
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)'''

    driver.close()
    

if __name__ == '__main__':
    run()
