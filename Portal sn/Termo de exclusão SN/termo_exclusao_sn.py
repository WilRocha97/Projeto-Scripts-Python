# -*- coding: utf-8 -*-
import os, time, shutil, re, fitz
from bs4 import BeautifulSoup
from requests import Session
from selenium import webdriver
from selenium.webdriver.common.by import By
from urllib3 import disable_warnings, exceptions
from pyautogui import prompt, confirm

from sys import path
path.append(r'..\..\_comum')
from sn_comum import _new_session_sn
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _download_file, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice


def renomear(empresa, download_folder, final_foder):
    cnpj, cpf, cod, nome = empresa
    
    count = 0
    while not len(os.listdir(download_folder)) > 0 :
        time.sleep(1)
        count += 1
        if count >= 30:
            print('❌ Erro, nenhum arquivo baixado')
            return False

    time.sleep(2)
    for file in os.listdir(download_folder):
        with fitz.open(os.path.join(download_folder, file)) as pdf:
            for page in pdf:
                textinho = page.get_text('text', flags=1 + 2 + 8)
                try:
                    data = re.compile(r'Data do fato motivador: (.+)').search(textinho).group(1)
                    nome = re.compile(r'Nome Empresarial: (.+)').search(textinho).group(1)
                except:
                    nome = re.compile(r'(.+)\nNome Empresarial:').search(textinho).group(1)
                
        file_rename = f'Termo de Exclusão do Simples Nacional {data.replace("/", "-")} - {cnpj} - {nome.replace("/", " ")}.pdf'
        file_rename = file_rename.replace('  ', ' ')
        shutil.move(os.path.join(download_folder, file), os.path.join(final_foder, file_rename))
        time.sleep(2)
    
    return True


def login(cnpj, cpf, cod, options):
    while True:
        # loga no site do simples nacional com web driver e retorna uma sessão de request
        status, driver = _initialize_chrome(options)
        driver, session = _new_session_sn(cnpj, cpf, cod, 'caixa postal', driver, usa_driver=True)
        
        if session == 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.':
            print('❌ Erro Login - Caracteres anti-robô inválidos. Tente novamente')
            driver.quit()
        else:
            return driver, session
            

def consulta(driver, empresa, comp, download_folder, final_foder):
    time.sleep(2)
    try:
        driver.get('https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATBHE/DTESN.app/')
    except:
        return 'erro'
    
    print('>>> Aguardando site')
    timer = 0
    while not _find_by_id('MenuPrincipal', driver):
        time.sleep(1)
        timer += 1
        if timer >= 10:
            return 'erro'

    try:
        termo = re.compile(r'TERMO DE EXCLUSÃO DO SIMPLES NACIONAL.+nº (.+), de (.+ de ' + comp + ')</a>').search(str(driver.page_source))
        numero = termo.group(1)
        data = termo.group(2)
    except:
        return f'✔ Não possuí termo de exclusão em {comp}'

    os.makedirs('ignore/Termos', exist_ok=True)
    os.makedirs('execução/Termos', exist_ok=True)
    
    link_mensagem = re.compile(r'href=\"(.+)\">TERMO DE EXCLUSÃO DO SIMPLES NACIONAL').search(str(driver.page_source)).group(1)
    driver.get(f'https://www8.receita.fazenda.gov.br{link_mensagem}')

    while not _find_by_id('MensagensSistema', driver):
        time.sleep(1)
    
    link_termo = re.compile(r'href=\"(.+)\">Acesso ao termo').search(str(driver.page_source)).group(1)
    driver.get(f'https://www8.receita.fazenda.gov.br{link_termo}')
    
    if not renomear(empresa, download_folder, final_foder):
        return '❌ Erro, nenhum arquivo baixado'
    
    return f'❗ Termo de exclusão do Simples Nacional;nº {numero};{data}'


@_time_execution
def run():
    comp = prompt(title='Script incrível', text='Qual o ano de emissão do termo de exclusão', default='0000')
    download_folder = "V:\\Setor Robô\\Scripts Python\\Portal sn\\Termo de exclusão SN\\ignore\\Termos"
    final_foder = "V:\\Setor Robô\\Scripts Python\\Portal sn\\Termo de exclusão SN\\execução\\Termos"
    
    if not comp:
        return False

    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cpf, cod, nome = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        # limpa a pasta de download
        for file in os.listdir(download_folder):
            os.remove(os.path.join(download_folder, file))
            
        while True:
            driver, session = login(cnpj, cpf, cod, options)
            # se existe uma sessão realiza a consulta
            if isinstance(session, Session):
                text = consulta(driver, empresa, comp, download_folder, final_foder)
                driver.quit()
                if text != 'erro':
                    break
            else:
                text = '❌ ' + session
                if session != 'Erro Login - Por favor, solucione o captcha para efetuar o acesso.':
                    driver.quit()
                    break
                print(text)
                driver.quit()
                
        # escreve na planilha de andamentos o resultado da execução atual
        print(text)
        _escreve_relatorio_csv(f'{cnpj};{cpf};{cod};{nome};{text[2:]}', nome='Termo de exclusão SN')

if __name__ == '__main__':
    run()
