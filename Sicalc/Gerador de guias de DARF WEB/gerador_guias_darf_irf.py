# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as p
import os, time, shutil

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome
from pyautogui_comum import _click_img, _wait_img


def login_sicalc(empresa, apuracao, vencimento, driver):
    cnpj, nome, valor, cod = empresa
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

    # esperar o botão de login habilitar e clicar nele
    _wait_img('continuar.png', conf=0.95, timeout=30)
    time.sleep(0.5)
    driver.find_element(by=By.XPATH, value='//*[@id="divBotoes"]/input[1]').click()

    # gerar a guia de DCTF
    gerar(empresa, apuracao, vencimento, driver)


def gerar(empresa, apuracao, vencimento, driver):
    cnpj, nome, valor, cod = empresa

    # esperar o menu referente a guia aparecer
    _wait_img('menu.png', conf=0.9, timeout=-1)
    time.sleep(1)

    print('>>> Inserindo informações da guia')
    # inserir o código da receita referente a guia
    driver.find_element(by=By.ID, value='codReceitaPrincipal').send_keys(cod)
    time.sleep(1)
    # confirmar a seleção
    p.press(['up', 'enter'], interval=1)
    p.moveTo(106, 375)
    p.click()
    time.sleep(1)

    # descer a visualização da página
    p.press('pgDn', presses=2)
    time.sleep(1)

    # clicar no campo e inserir a data de apuração
    _click_img('apuracao.png', conf=0.95)
    time.sleep(1)
    p.write(apuracao)
    time.sleep(1)

    # clicar no campo e inserir o valor
    _click_img('valor.png', conf=0.95, clicks=2)
    time.sleep(1)
    p.write(valor)

    p.press('tab')
    time.sleep(1)

    # clica em calcular
    driver.find_element(by=By.ID, value='btnCalcular').click()

    # espera guia ser calculada
    _wait_img('checkbox_guia.png', conf=0.9, timeout=-1)

    print('>>> Guia calculada')
    # marca a guia que será gerada
    driver.find_element(by=By.XPATH, value='//*[@id="cts"]/tbody/tr/td[1]/input').click()
    time.sleep(1)
    
    print('>>> Gerando guia')
    driver.find_element(by=By.ID, value='btnDarf').click()
    
    while not os.path.exists('V:/Setor Robô/Scripts Python/Sicalc/Gerador de guias de DARF WEB/execucao/Guias/Darf.pdf'):
        time.sleep(1)
    while os.path.exists('V:/Setor Robô/Scripts Python/Sicalc/Gerador de guias de DARF WEB/execucao/Guias/Darf.pdf'):
        try:
            arquivo = f'{nome} - {cnpj} - DARF IRRF - {apuracao.replace("/", "-")} - venc. {vencimento.replace("/", "-")}.pdf'
            download_folder = "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execucao\\Guias"
            shutil.move(os.path.join(download_folder, 'Darf.pdf'), os.path.join(download_folder, arquivo))
            time.sleep(2)
        except:
            pass

    print('✔ Guia gerada')
    _escreve_relatorio_csv('{};{};{};{};Guia gerada'.format(cnpj, nome, valor, cod))
    
    driver.close()


@_time_execution
def run():
    os.makedirs('execucao/Guias', exist_ok=True)
    # p.mouseInfo()
    apuracao = p.prompt(title='Script incrível', text='Qual o período de apuração?', default='00/0000')
    vencimento = p.prompt(title='Script incrível', text='Qual o vencimento?\n(Sempre dia 20, se for fim de semana ou feriado, antecipa.)', default='00/00/0000')
    
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
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execucao\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)

        # iniciar o driver do chome
        status, driver = _initialize_chrome(options)

        # fazer login do SICALC
        login_sicalc(empresa, str(apuracao), vencimento, driver)


if __name__ == '__main__':
    run()
