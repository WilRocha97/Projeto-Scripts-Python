# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import os, time, re, csv, shutil, fitz, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome
from pyautogui_comum import _click_img, _find_img, _wait_img


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

    dados = "V:\\Setor Robô\\Scripts Python\\_comum\\DadosVeri-SIEG.txt"
    f = open(dados, 'r', encoding='utf-8')
    user = f.read()
    user = user.split('/')
    
    # inserir o email no campo
    driver.find_element(by=By.ID, value='txtEmail').send_keys(user[0])
    time.sleep(1)
    
    # inserir a senha no campo
    driver.find_element(by=By.ID, value='txtPassword').send_keys(user[1])
    time.sleep(1)
    
    # clica em acessar
    driver.find_element(by=By.ID, value='btnSubmit').click()
    time.sleep(1)
    
    return driver
    

def sieg_iris(driver, modulo):
    print('>>> Acessando IriS DCTF WEB')
    if modulo == 'Termo(s) de Exclusão do Simples Nacional':
        driver.get('https://hub.sieg.com/IriS/#/Mensagens?Filtro=Exclusoes')
    if modulo == 'Termos(s) de Intimação':
        driver.get('https://hub.sieg.com/IriS/#/Mensagens?Filtro=Intimacoes')
    
    return driver


def imprime_mensagem(driver):
    # aguarda e clica na mensagem
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]'):
        time.sleep(1)
        
    # aguarda e clica para filtrar abrir o dropdown de filtro de mensagens não lidas
    while not localiza_id(driver, 'select2-ddlSelectedMessage-container'):
        time.sleep(1)
    
    _click_img('filtro_lidas.png', conf=0.9)
    _wait_img('opcao_filtro_n_lidas.png', conf=0.9)
    _click_img('opcao_filtro_n_lidas.png', conf=0.9)
    time.sleep(1)
    
    # aguarda e clica na mensagem
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]'):
        time.sleep(1)
    
    return driver
    

def salvar_pdf(driver, pasta_analise):
    driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]').click()
    
    # aguarda a janela de pré-visualização abrir e clica em imprimir
    _wait_img('imprimir.png', conf=0.9)
    _click_img('imprimir.png', conf=0.9)
    
    # aguarda a tela de impressão
    _wait_img('tela_imprimir.png', conf=0.9)
    
    # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
    if _find_img('print_to_pdf.png', conf=0.9):
        _click_img('print_to_pdf.png', conf=0.9)
        # aguarda aparecer a opção de salvar como PDF e clica nela
        _wait_img('salvar_como_pdf.png', conf=0.9)
        _click_img('salvar_como_pdf.png', conf=0.9)
    
    # aguarda aparecer o botão de salvar e clica nele
    _wait_img('botao_salvar.png', conf=0.9)
    _click_img('botao_salvar.png', conf=0.9)
    
    print('>>> Salvando relatório')
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
    
    os.makedirs(pasta_analise, exist_ok=True)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    erro = 'sim'
    while erro == 'sim':
        try:
            pyperclip.copy(pasta_analise)
            p.hotkey('ctrl', 'v')
            erro = 'não'
        except:
            erro = 'sim'
    
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(1)
    p.hotkey('ctrl', 'w')
    
    print('✔ Mensagem salva com sucesso')
    return 'Mensagem salva com sucesso'


def verifica_mesagem(pasta_analise, pasta_final):
    print('>>> Analisando mensagem')
    # Analisa cada pdf que estiver na pasta
    for arquivo in os.listdir(pasta_analise):
        # Abrir o pdf
        arq = os.path.join(pasta_analise, arquivo)
        
        with fitz.open(arq) as pdf:
            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                print(textinho)
                time.sleep(22)
        
        os.makedirs(pasta_final, exist_ok=True)
        shutil.move(arq, pasta_final)
        
    print(situacao.replace(';', ' | '))
    return situacao


@_time_execution
def run():
    modulo = p.confirm(title='Script incrível', buttons=('Termo(s) de Exclusão do Simples Nacional', 'Termos(s) de Intimação'))
    pasta_analise = r'V:\Setor Robô\Scripts Python\SIEG\Download Menságens Importantes\ignore\Mensagens'
    pasta_final = os.path.join(r'V:\Setor Robô\Scripts Python\SIEG\Download Menságens Importantes\execução', modulo)
    
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.9")
    
    # iniciar o driver do chrome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)
    driver = sieg_iris(driver, modulo)
    
    if imprime_mensagem(driver):
        salvar_pdf(driver, pasta_analise)
        verifica_mesagem(pasta_analise, pasta_final)
        
    driver.close()
    

if __name__ == '__main__':
    run()
