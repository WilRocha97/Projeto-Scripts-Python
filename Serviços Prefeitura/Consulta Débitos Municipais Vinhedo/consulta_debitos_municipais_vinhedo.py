# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
from sys import path
from selenium.webdriver.support.select import Select
import os, time, re
from pyautogui import press, hotkey, write, click, alert, prompt, getWindowsWithTitle

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha
from pyautogui_comum import _find_img, _wait_img, _click_img

os.makedirs('execucao/documentos', exist_ok=True)


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
    

def login(insc_muni):
    status, driver = initialize_chrome()
    
    url_inicio = 'http://vinhedomun.presconinformatica.com.br/certidaoNegativa.jsf?faces-redirect=true'
    driver.get(url_inicio)

    while not find_by_id('homeForm:panelCaptcha:j_idt47', driver):
        time.sleep(1)
    # bloco que salva a iamgem do captcha
    element = driver.find_element(by=By.ID, value='homeForm:panelCaptcha:j_idt47')
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
    # resolve o captcha
    captcha = _solve_text_captcha(os.path.join('ignore', 'captcha', 'captcha.png'))
    
    # espera o campo do tipo da pesquisa
    while not find_by_id('homeForm:inputTipoInscricao_label', driver):
        time.sleep(1)
    # clica no menu
    driver.find_element(by=By.ID, value='homeForm:inputTipoInscricao_label').click()
    
    # espera o menu abrir
    while not find_by_path('/html/body/div[6]/div/ul/li[2]', driver):
        time.sleep(1)
    # clica na opção "Mobiliário"
    driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/ul/li[2]').click()
    
    if not captcha:
        print('Erro Login - não encontrou captcha')
    
    # clica na barra de inscrição e insere
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').click()
    time.sleep(2)
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').send_keys(insc_muni)
    _send_input('homeForm:panelCaptcha:j_idt50', captcha, driver)
    time.sleep(2)
    
    # clica no botão de pesquisar
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[2]/div[5]/a/div').click()

    print('>>> Consultando')
    while 'dados-contribuinte-inscricao">0000000000' not in driver.page_source:
        if find_by_path('/html/body/div[1]/div[5]/div[2]/div[1]/div/ul/li/span', driver):
            padraozinho = re.compile(r'<span class="ui-messages-error-summary">(.+)<\/span><\/li><\/ul><\/div><\/div><div class="right"><button id="j_idt70"')
            situacao = padraozinho.search(driver.page_source).group(1)
            situacao_print = f'❌ {situacao}'
            driver.quit()
            return situacao, situacao_print
        time.sleep(1)
    
    situacao_print = f'❌ {situacao}'
    driver.quit()
    return situacao, situacao_print
    
    
def salvar_guia():
    # Salvar o PDF
    while not _find_img('SalvarPDF.png', conf=0.9):
        sleep(1)
        
    # Usa o pyperclip porque o pyautogui não digita letra com acento
    copy(texto)
    hotkey('ctrl', 'v')
    sleep(0.5)
    
    # Selecionar local
    press('tab', presses=6)
    sleep(0.5)
    press('enter')
    sleep(0.5)
    
    copy('V:\Setor Robô\Scripts Python\CUCA\Relatórios DP-CUCA\{}\{}'.format(e_dir, diretorio))
    
    hotkey('ctrl', 'v')
    sleep(0.5)
    press('enter')
    
    sleep(0.5)
    hotkey('alt', 'l')
    sleep(1)
    if _find_img('SalvarComo.png', 0.9):
        press('s')
        sleep(1)
    

@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, insc_muni, cidade = empresa
        
        _indice(count, total_empresas, empresa)
        
        situacao, situacao_print = login(insc_muni)
        print(situacao_print)
        if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
            return False
        
        _escreve_relatorio_csv(f'{cnpj};{insc_muni};{situacao}', nome='Consulta de débitos municipais de Vinhedo')

    return True
    
    
if __name__ == '__main__':
    run()
