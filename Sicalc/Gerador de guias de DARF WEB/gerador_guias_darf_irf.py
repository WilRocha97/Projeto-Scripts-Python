# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui as p
import os, time, shutil

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome
from pyautogui_comum import _click_img, _wait_img, _find_img


def renomear(empresa, apuracao, vencimento):
    cnpj, nome, nota, valor, cod = empresa
    download_folder = "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias"
    guia = os.path.join(download_folder, 'Darf.pdf')
    while os.path.exists(guia):
        try:
            arquivo = f'{nome.replace("/", " ")} - {cnpj} - DARF IRRF {cod} {apuracao.replace("/", "-")} - venc. {vencimento.replace("/", "-")}.pdf'
            shutil.move(guia, os.path.join(download_folder, arquivo))
            time.sleep(2)
        except:
            pass


def login_sicalc(empresa, apuracao, vencimento, driver):
    cnpj, nome, nota, valor, cod = empresa
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
    timer = 0
    while not _find_img('continuar.png', conf=0.95):
        time.sleep(1)
        timer += 1
        if timer >= 10:
            return False
        
    time.sleep(0.5)
    driver.find_element(by=By.XPATH, value='//*[@id="divBotoes"]/input[1]').click()

    # gerar a guia de DCTF
    if not gerar(empresa, apuracao, vencimento, driver):
        return False
    time.sleep(2)
    return True


def gerar(empresa, apuracao, vencimento, driver):
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
        driver.find_element(by=By.ID, value='observacao').send_keys(f'Referênte à NF: {nota} - R. POSTAL')
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
    while not _find_img('checkbox_guia.png', conf=0.95):
        driver.find_element(by=By.ID, value='btnCalcular').click()
        # descer a visualização da página
        p.press('pgDn')
        time.sleep(1)
        
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
        
    renomear(empresa, apuracao, vencimento)

    print('✔ Guia gerada')
    _escreve_relatorio_csv('{};{};{};{};Guia gerada'.format(cnpj, nome, valor, cod))

    driver.close()
    while _find_img('erro_site.png', conf=0.9):
        driver.close()
        time.sleep(2)
    return True


@_time_execution
def run():
    os.makedirs('execução/Guias', exist_ok=True)
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
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        
        erro = 'sim'
        while erro == 'sim':
            try:
                # iniciar o driver do chome
                status, driver = _initialize_chrome(options)
        
                # fazer login do SICALC
                if not login_sicalc(empresa, str(apuracao), vencimento, driver):
                    driver.close()
                    while _find_img('erro_site.png', conf=0.9):
                        driver.close()
                        time.sleep(1)
                    erro = 'sim'
                else:
                    erro = 'nao'
            except:
                erro = 'sim'


if __name__ == '__main__':
    run()
