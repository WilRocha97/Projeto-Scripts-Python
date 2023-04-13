# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, shutil, re, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from pyautogui_comum import _click_img, _wait_img, _find_img
from captcha_comum import _solve_hcaptcha
from chrome_comum import _initialize_chrome


def renomear(driver, empresa, apuracao, tipo):
    cnpj, nome, nota, valor, cod = empresa
    
    vencimento = re.compile(r'data-darf=\"true\".+(\d\d/\d\d/\d\d\d\d)</td><td>\d\d/\d\d/').search(driver.page_source).group(1)
    
    download_folder = "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias"
    guia = os.path.join(download_folder, 'Darf.pdf')
    while os.path.exists(guia):
        try:
            arquivo = f'{nome.replace("/", " ")} - {cnpj} - {tipo} {cod} {apuracao.replace("/", "-")} - venc. {vencimento.replace("/", "-")}.pdf'
            shutil.move(guia, os.path.join(download_folder, arquivo))
            time.sleep(2)
        except:
            pass


def login_sicalc(empresa):
    cnpj, nome, nota, valor, cod = empresa
    
    cnpj_formatado = f'{cnpj[0:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}'
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36'}
    
    with Session() as s:
        '''erro = 'sim'
        while erro == 'sim':
            try:'''
        s.get('https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte', headers=headers)
        hcaptcha_data = {'url': 'https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte', 'sitekey': '20a82aa0-d5ea-4ae2-8b4c-ac341dfe6ee7'}
        token = _solve_hcaptcha(hcaptcha_data)
        token = str(token)

        payload = {'cnpjFormatado': cnpj_formatado,
                   'h-captcha-response': token}
        s.post('https://sicalc.receita.economia.gov.br/sicalc/api/captcha', headers=headers, data=payload)
        
        res = s.get('https://sicalc.receita.economia.gov.br/sicalc/rapido/calculo', headers=headers)
        soup = BeautifulSoup(res.content, 'html.parser')
        soup = soup.prettify()
        erro = 'não'
        '''except:
            erro = 'sim'''
            
        # opções para fazer com que o chome trabalhe em segundo plano (opcional)
        options = webdriver.ChromeOptions()
        # options.add_argument('--headless')
        # options.add_argument('--window-size=1920,1080')
        options.add_argument("--start-maximized")
        options.add_experimental_option('prefs', {
            "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias",  # Change default directory for downloads
            "download.prompt_for_download": False,  # To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
        })
        status, driver = _initialize_chrome(options)
        for cookie in s.get.cookies():
            driver.cookies.set(cookie['name'], cookie['value'])
        
        driver.get('https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte')
        print('>>> Acessando o site')
        time.sleep(1)
        
        # clicar na opção de CNPJ e inserir o CNPJ no campo
        driver.find_element(by=By.ID, value='optionPJ').click()
        time.sleep(1)
        driver.find_element(by=By.ID, value='cnpjFormatado').send_keys(cnpj)
        time.sleep(1)
        
        # selecionar o checkbox do captcha
        _click_img('checkbox.png', conf=0.95)
        p.moveTo(5, 5)
        p.press('pgDn')
        
        time.sleep(4)
    return True


def gerar(empresa, apuracao, driver, tipo):
    cnpj, nome, nota, valor, cod = empresa

    # esperar o menu referente a guia aparecer
    timer = 0
    while not _find_img('menu.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer >= 10:
            return False
    
    print('>>> Inserindo informações da guia')
    
    if nota.split(','):
        # insere observações na guia
        driver.find_element(by=By.ID, value='observacao').send_keys(f'Referente à NF: {nota} - R. POSTAL')
        time.sleep(1)
    
    # inserir o código da receita referente a guia
    driver.find_element(by=By.ID, value='codReceitaPrincipal').send_keys(cod)
    time.sleep(1)
    
    # confirmar a seleção
    p.press(['up', 'enter'], interval=1)
    p.moveTo(106, 375)
    p.click()
    time.sleep(1)

    # descer a visualização da página
    timer = 0
    while not _find_img('apuracao.png', conf=0.95):
        p.press('pgDn')
        time.sleep(1)
        timer += 1
        if timer >= 20:
            return False
        
    # clicar no campo e inserir a data de apuração
    _click_img('apuracao.png', conf=0.95)
    time.sleep(1)
    p.write(apuracao)
    time.sleep(1)

    # descer a visualização da página
    while not _find_img('valor.png', conf=0.95):
        p.press('pgDn')
        time.sleep(1)
        
    # clicar no campo e inserir o valor
    _click_img('valor.png', conf=0.95, clicks=2)
    time.sleep(1)
    p.write(valor)

    p.press('tab')
    time.sleep(1)

    # espera guia ser calculada
    timer = 0
    while not _find_img('checkbox_guia.png', conf=0.95):
        driver.find_element(by=By.ID, value='btnCalcular').click()
        # descer a visualização da página
        p.press('pgDn')
        time.sleep(1)
        timer += 1
        if timer >= 20:
            return False
        
    print('>>> Guia calculada')
    time.sleep(1)
    
    # marca a guia que será gerada
    erro = 'sim'
    timer = 0
    while erro == 'sim':
        try:
            driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[2]/div[1]/div/div/div/div[4]/table/tbody/tr/td[1]/input').click()
            time.sleep(1)
            erro = 'não'
        except:
            driver.find_element(by=By.ID, value='btnCalcular').click()
            erro = 'sim'
        
        time.sleep(1)
        timer+= 1
        if timer >= 15:
            return False

    print('>>> Gerando guia')
    download_folder = "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias"
    guia = os.path.join(download_folder, 'Darf.pdf')
    while not os.path.exists(guia):
        driver.find_element(by=By.ID, value='btnDarf').click()
        time.sleep(5)
        p.hotkey('Ctrl', 'Shift', 'Tab')
        
    renomear(driver, empresa, apuracao, tipo)

    print('✔ Guia gerada')
    _escreve_relatorio_csv('{};{};{};{};Guia gerada'.format(cnpj, nome, valor, cod), nome=f'Resumo gerador {tipo}')

    driver.close()
    return True


@_time_execution
def run():
    os.makedirs('execução/Guias', exist_ok=True)
    # p.mouseInfo()
    tipo = p.confirm(title='Script incrível', text='Qual tipo da guia?', buttons=('DARF IRRF', 'DARF DP'))
    apuracao = p.prompt(title='Script incrível', text='Qual o período de apuração?', default='00/0000')
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        
        login_sicalc(empresa)
        

if __name__ == '__main__':
    run()