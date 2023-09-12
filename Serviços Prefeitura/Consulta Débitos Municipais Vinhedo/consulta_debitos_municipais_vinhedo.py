# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import os, time, re
from pyautogui import press, hotkey
from pyperclip import copy

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_path, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha
from pyautogui_comum import _find_img, _wait_img, _click_img

os.makedirs('execução/Certidões', exist_ok=True)


def pesquisar(options, cnpj, insc_muni):
    status, driver = _initialize_chrome(options)
    
    url_inicio = 'http://vinhedomun.presconinformatica.com.br/certidaoNegativa.jsf?faces-redirect=true'
    driver.get(url_inicio)
    
    contador = 1
    while not _find_by_id(f'homeForm:panelCaptcha:j_idt{str(contador)}', driver):
        contador += 1
        time.sleep(0.2)
        
    # resolve o captcha
    captcha = _solve_text_captcha(driver, 'homeForm:panelCaptcha:j_idt47')
    
    # espera o campo do tipo da pesquisa
    while not _find_by_id('homeForm:inputTipoInscricao_label', driver):
        time.sleep(1)
    # clica no menu
    driver.find_element(by=By.ID, value='homeForm:inputTipoInscricao_label').click()
    
    # espera o menu abrir
    while not _find_by_path('/html/body/div[6]/div/ul/li[2]', driver):
        time.sleep(1)
    # clica na opção "Mobiliário"
    driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/ul/li[2]').click()
    
    if not captcha:
        print('Erro Login - não encontrou captcha')

    time.sleep(1)
    # clica na barra de inscrição e insere
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').click()
    time.sleep(2)
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').send_keys(insc_muni)
    _send_input(f'homeForm:panelCaptcha:j_idt{str(contador + 3)}', captcha, driver)
    time.sleep(2)
    
    # clica no botão de pesquisar
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[2]/div[5]/a/div').click()

    print('>>> Consultando')
    while 'dados-contribuinte-inscricao">0000000000' not in driver.page_source:
        if _find_by_path('/html/body/div[1]/div[5]/div[2]/div[1]/div/ul/li/span', driver):
            padraozinho = re.compile(r'confirmationMessage\" class=\"ui-messages ui-widget\" aria-live=\"polite\"><div '
                                     r'class=\"ui-messages-error ui-corner-all\"><span class=\"ui-messages-error-icon\"></span><ul><li><span '
                                     r'class=\"ui-messages-error-summary\">(.+).</span></li></ul></div></div><div class=\"right\"><button id=\"j_idt')
            situacao = padraozinho.search(driver.page_source).group(1)
            situacao_print = f'❌ {situacao}'
            # print(driver.page_source)
            driver.quit()
            return situacao, situacao_print
        
        time.sleep(1)

    situacao = salvar_guia(cnpj)
    situacao_print = f'✔ {situacao}'
    driver.quit()
    return situacao, situacao_print
    
    
def salvar_guia(cnpj):
    print('>>> Salvando Certidão Negativa')
    # Salvar o PDF
    while not _find_img('salvar_pdf.png', conf=0.9):
        time.sleep(1)
    _click_img('salvar_pdf.png', conf=0.9)
    
    while not _find_img('salvar.png', conf=0.9):
        time.sleep(1)
    _click_img('salvar.png', conf=0.9)
    
    while not _find_img('tela_impressao.png', conf=0.9):
        time.sleep(1)
        
    # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
    imagens = ['print_to_pdf.png', 'print_to_pdf_2.png']
    for img in imagens:
        if _find_img(img, conf=0.9) or _find_img(img, conf=0.9):
            _click_img(img, conf=0.9)
            # aguarda aparecer a opção de salvar como PDF e clica nela
            _wait_img('salvar_como_pdf.png', conf=0.9)
            _click_img('salvar_como_pdf.png', conf=0.9)
    
    # aguarda aparecer o botão de salvar e clica nele
    _wait_img('botao_salvar.png', conf=0.9)
    _click_img('botao_salvar.png', conf=0.9)
    
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
        
    # Usa o pyperclip porque o pyautogui não digita letra com acento
    while True:
        try:
            copy(f'{cnpj} - Certidão Negativa de Débitos Municipais Vinhedo.pdf')
            hotkey('ctrl', 'v')
            break
        except:
            pass
    
    # Selecionar local
    press('tab', presses=6)
    time.sleep(0.5)
    press('enter')
    time.sleep(0.5)
    
    # Usa o pyperclip porque o pyautogui não digita letra com acento
    while True:
        try:
            copy('V:\Setor Robô\Scripts Python\Serviços Prefeitura\Consulta Débitos Municipais Vinhedo\execução\Certidões')
            hotkey('ctrl', 'v')
            break
        except:
            pass
    
    time.sleep(0.5)
    hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('salvar_como.png', conf=0.9):
        press('s')
        time.sleep(1)
    
    return 'Certidão negativa salva'


@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, insc_muni, nome = empresa
        
        _indice(count, total_empresas, empresa)
        
        situacao, situacao_print = pesquisar(options, cnpj, insc_muni)
        print(situacao_print)
        if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
            return False
        
        _escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{situacao}', nome='Consulta de débitos municipais de Vinhedo')

    return True
    
    
if __name__ == '__main__':
    run()
