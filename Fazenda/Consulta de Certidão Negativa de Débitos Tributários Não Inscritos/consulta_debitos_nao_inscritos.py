# -*- coding: utf-8 -*-
import time, re, os, fitz, shutil
from requests.exceptions import ConnectionError
from bs4 import BeautifulSoup
from sys import path
from requests import Session
from time import sleep
from selenium import webdriver
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By

path.append(r'..\..\_comum')
from fazenda_comum import _get_info_post, _new_session_fazenda_driver
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _download_file, _open_lista_dados, _where_to_start, _indice
from chrome_comum import _find_by_id, _find_by_path


def abre_pagina_consulta(driver):
    print('>>> Abrindo Conta Fiscal')

    try:
        while not re.compile(r'>Conta Fiscal do ICMS e Parcelamento').search(driver.page_source):
            try:
                button = driver.find_element(by=By.XPATH,
                                             value='/html/body/div[2]/section/div/div/div/div[2]/div/ul/li/form/div[5]/div/a')
                button.click()
                time.sleep(3)
            except:
                pass
            time.sleep(1)
    except:
        print('❗ Erro ao logar na empresa, tentando novamente')
        return driver, 'erro'

    url_consulta = re.compile(r'<a href=\"(.+\d).+>Conta Fiscal do ICMS e Parcelamento').search(
        driver.page_source).group(1)

    driver.get(url_consulta)

    while not _find_by_id('divcontainer', driver):
        time.sleep(1)

    print('>>> Abrindo consulta de CNDNI')
    while True:
        try:
            url_consulta_cndni = re.compile(
                r'href="(https:\/\/www10.fazenda.sp.gov.br\/CertidaoNegativaDeb\/Pages.+)" tabindex="-1">Verificar Impedimentos eCND').search(
                driver.page_source).group(1)
            break
        except:
            pass

    driver.get(url_consulta_cndni)

    return driver, 'ok'


def consulta_cndni(driver, nome, cnpj, pasta_inicial):
    contador = 0
    while True:
        print('>>> Consultando.')
        while not _find_by_id('MainContent_txtDocumento', driver):
            time.sleep(1)


        time.sleep(1)
        driver.find_element(by=By.ID, value='MainContent_txtDocumento').clear()
        time.sleep(1)
        driver.find_element(by=By.ID, value='MainContent_txtDocumento').send_keys(cnpj)
        time.sleep(1)
        driver.find_element(by=By.ID, value='MainContent_btnPesquisar').click()

        # Wait for the alert to be displayed and store it in a variable
        try:
            wait = WebDriverWait(driver, 5)
            alert = wait.until(expected_conditions.alert_is_present())
            if alert:
                # Store the alert text in a variable
                text = alert.text
                # Press the OK button
                alert.accept()

                if text != 'Pesquisa não autorizada. Cadastro não localizado.':
                    print(f'❌ {text}')
                    return driver, text
                else:
                    print(f'❗ Possível erro ao digitar o CNPJ, tentando novamente.')
                    contador += 1

                if contador >= 3:
                    print(f'❌ {text}')
                    return driver, text

            else:
                break
        except:
            break

    contador = 0
    while not _find_by_id('MainContent_lnkImprimirCertidaoBotao1', driver):
        print('>>> Consultando..')
        time.sleep(1)
        contador += 1
        if contador > 15:
            print('❌ Erro ao consultar CNDNI, tentando novamente')
            return driver, 'erro'

    contador = 0
    while True:
        if re.compile(r"Server Error in '/CertidaoNegativaDeb' Application.").search(driver.page_source):
            print('❌ Erro ao acessar o site, tentando novamente')
            return driver, 'erro'
        
        try:
            driver.find_element(by=By.ID, value='MainContent_lnkImprimirCertidaoBotao1').click()
            resultado = renomeia_cndni(nome, cnpj, pasta_inicial)
            if resultado:
                break
            
        except:
            if re.compile(r"Ocorreu uma falha na geração do relatório!").search(driver.page_source):
                driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[3]/div/button').click()
                print('❌ Erro ao emitir relatório, tentando novamente')
                return driver, 'erro'
                
            driver.execute_script("window.scrollBy(0,200)")
            time.sleep(0.5)
            print('>>> Consultando...')
            contador += 1
        
        if re.compile(r'Acesso negado').search(driver.page_source):
            print('❌ Acesso negado para essa empresa')
            return driver, 'Acesso negado para essa empresa'
        
        if contador > 60:
            print('❌ Erro ao consultar CNDNI, tentando novamente')
            return driver, 'erro'
    
    driver.execute_script('document.getElementById("MainContent_lnkVoltar").click()')
    return driver, resultado


def renomeia_cndni(nome, cnpj, pasta_inicial):
    while os.listdir(pasta_inicial) == []:
        time.sleep(1)
    
    debitos = 'não'
    time.sleep(1)
    for cndni in os.listdir(pasta_inicial):
        arq = os.path.join(pasta_inicial, cndni)
        if arq.endswith('.crdownload'):
            os.remove(arq)
            return False
            
        while True:
            try:
                doc = fitz.open(arq, filetype="pdf")
                break
            except:
                print('>>> Aguardando download')
                pass
        
        for page in doc:
            texto = page.get_text('text', flags=1 + 2 + 8)
            if re.compile(r'\nHá Pendências').search(texto):
                debitos = 'sim'
            if re.compile(r'\nHá Débitos').search(texto):
                debitos = 'sim'
            
            if debitos == 'sim':
                debitos_lista = ''
                
                resumo = [('ICMS Declarado', r'ICMS Declarado\nNão há Débitos\n'),
                          ('ICMS Parcelamento', r'ICMS Parcelamento\nNão há Débitos\n'),
                          ('IPVA', r'IPVA\nNão há Débitos\n'),
                          ('ITCMD', r'ITCMD\nNão há Débitos\n'),
                          ('AIIM', r'AIIM\nNão há Débitos\n'),
                          ('ICMS Pendência', r'ICMS Pendência\nNão há Pendências\n')]
                for item in resumo:
                    if not re.compile(item[1]).search(texto):
                        debitos_lista += ' - ' + item[0]
                
                resumo = [('GIA', r'\nGIA\n'),
                          ('GIA-EFD', r'\nGIA\/EFD\n'),]
                for item in resumo:
                    if re.compile(item[1]).search(texto):
                        debitos_lista += ' - ' + item[0]
                
                if debitos_lista != '':
                    doc.close()
                    pasta_debito = os.path.join('Execução', 'CNDNI com débitos')
                    os.makedirs(pasta_debito, exist_ok=True)
                    shutil.move(arq, os.path.join(pasta_debito, f'{nome[:30]} - {cnpj} - CNDNI Débitos{debitos_lista}.pdf'))
                    print('❗ Com Débitos')
                    return 'Com débitos'
        
        doc.close()
        pasta_sem_debito = os.path.join('Execução', 'CNDNI')
        os.makedirs(pasta_sem_debito, exist_ok=True)
        shutil.move(arq, os.path.join(pasta_sem_debito, f'{nome[:30]} - {cnpj} - CNDNI.pdf'))
    
    print('✔ Sem débitos')
    return 'Sem débitos'


@_time_execution
def run():
    andamentos = 'Consulta de Débitos Não Inscritos'
    pasta_inicial = "V:\\Setor Robô\\Scripts Python\\Fazenda\\Consulta de Certidão Negativa de Débitos Tributários Não Inscritos\\ignore\\CNDNI"
    
    for cndni in os.listdir(pasta_inicial):
        os.remove(os.path.join(pasta_inicial, cndni))
        
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": pasta_inicial,  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    driver = ''
    for count, empresa in enumerate(empresas[index:], start=1):
        codigo, cnpj, nome, usuario, senha, perfil = empresa
        nome = nome.replace('/', '')
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        resultado = 'ok'
        while True:
            # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
            if usuario_anterior != usuario:
                # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
                try:
                    driver.close()
                except:
                    pass
                
                contador = 0
                while True:
                    try:
                        driver, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil, retorna_driver=True, options=options)
                        break
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
                    
                    if contador >= 5:
                        print('❌ Impossível de logar com esse usuário')
                        sid = 'Impossível de logar com esse usuário'
                        driver = 'erro'
                        break
                    
                if driver == 'erro':
                    _escreve_relatorio_csv(f'{codigo};{cnpj};{nome};{sid}', nome=andamentos)
                    usuario_anterior = usuario
                    break
                    
                else:
                    driver, resultado = abre_pagina_consulta(driver)
            
            if resultado == 'ok':
                driver, resultado = consulta_cndni(driver, nome, cnpj, pasta_inicial)
                
                if resultado != 'erro':
                    _escreve_relatorio_csv(f'{codigo};{cnpj};{nome};{resultado}', nome=andamentos)
                    usuario_anterior = usuario
                    break
                    
                driver.close()
                usuario_anterior = 'padrão'
                continue
    
    try:
        driver.close()
    except:
        pass
    return True


if __name__ == '__main__':
    run()
    