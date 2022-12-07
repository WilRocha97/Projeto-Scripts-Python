# -*- coding: utf-8 -*-
import re
import time

from bs4 import BeautifulSoup
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _download_file, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice, _escreve_header_csv
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome, _find_by_id

_site_key = '6LcUdlwaAAAAADjMUnV5jwUKKMlqdZbKNKXTfYEi'
_execucao = 'Valor FAP original'


def login(cnpj, options):
    tentativas = 1
    x = 'erro'
    while x == 'erro':
        status, driver = _initialize_chrome(options)
        url = 'https://www2.dataprev.gov.br/FapWeb/pages/login.xhtml'
        print('>>> Acessando o site')
        try:
            driver.get(url)
            x = 'ok'
        except:
            driver.close()
            time.sleep(5)
            print('>>> Erro no site, tentando novamente')
            x = 'erro'
        
        if x == 'ok':
            sleep(5)
            try:
                recaptcha_data = {'sitekey': _site_key, 'url': url}
        
                captcha = _solve_recaptcha(recaptcha_data)
                print('>>> logando no site')
                driver.find_element(by=By.ID, value='form:CNPJ').send_keys(cnpj)
                driver.find_element(by=By.ID, value='form:senha').send_keys(str(cnpj[:8]))
                driver.execute_script('document.getElementById("g-recaptcha-response").innerHTML="' + captcha + '";')
        
                sleep(1)
                driver.find_element(by=By.ID, value='form:btnAutenticar').click()
                sleep(5)
        
                html = driver.page_source.encode('utf-8')
                soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')
                padrao = re.compile(r'<li class=\"erro\">(.+)</li>')
                resposta = padrao.search(str(soup))
                if resposta:
                    resposta = resposta.group(1)
                    resposta = resposta.replace('O parâmetro responde é inválido.</li><li class="erro">', '')
                    print('❌ {}'.format(resposta))
                    
                    if resposta == 'Senha da empresa inválida! ' \
                            or resposta == 'Sem acesso: Verifique senha, CNPJ da empresa ou nível de acesso!' \
                            or resposta == 'Irregularidade com a senha digitada. Procure a unidade de atendimento da Receita Federal do Brasil - RFB ou consulte o Guia de Consulta na página inicial do FAP. ':
                        _escreve_relatorio_csv(';'.join([cnpj, resposta]), nome=_execucao)
                        driver.close()
                        return 'erro no login'
                    
                    elif resposta == 'O parâmetro responde é inválido.</li><li class="erro">Erro ao validar o captcha.':
                        x = 'erro'
                        driver.close()
                        
                    else:
                        tentativas += 1
                        x = 'erro'
                        driver.close()
                    
                else:
                    resposta = consulta(driver, cnpj)
                    
                    if resposta == 'erro':
                        x = 'erro'
                        tentativas += 1
                        driver.close()
                        
                    else:
                        driver.close()
                        return 'ok'
            except:
                resposta = 'Erro no login'
                print('❌ {}'.format(resposta))
                tentativas += 1
                x = 'erro'
                driver.close()
    
            if tentativas >= 3:
                _escreve_relatorio_csv(';'.join([cnpj, resposta]), nome=_execucao)
                return 'erro no login'


def consulta(driver, cnpj):
    print('>>> consultando')
    try:
        driver.get('https://www2.dataprev.gov.br/FapWeb/pages/consulta/resultadoConsultaFap.xhtml')
    except:
        return 'erro'
    while not _find_by_id('form:panelConsulta', driver):
        sleep(1)

    sleep(1)

    # selecionar uma competência
    # driver.execute_script('document.getElementById("form:anoVigencia").value=2020')
    # driver.execute_script('document.getElementById("form:btnConsultaCnpjDigitado").click(2)')
    # sleep(5)

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser', from_encoding='utf-8')

    try:
        padrao = re.compile(r'(Valor do FAP Original)(\n.+){8}\">(.+)(\n.+){5}\">(.+)')
        vigencia = re.compile(r'Ano de Vigência: .+\n.+(\d\d\d\d)').search(str(soup)).group(1)
        resposta = padrao.search(str(soup))
        valor = resposta.group(1) + ': ' + resposta.group(3)
        data = resposta.group(5)

        print(f'✔ {valor} | Data do cálculo: {data} | Vigência: {vigencia}')
        _escreve_relatorio_csv(f'{cnpj};{valor};{data};{vigencia}', nome=_execucao)

    except:
        try:
            padrao = re.compile(r'<div class=\"mensagem\"><ul><li class=\"erro\">(.+).</li><li')
            resposta = padrao.search(str(soup))
            erro = resposta.group(1)

            print('❌ {}'.format(erro))
            _escreve_relatorio_csv(f'{cnpj};{erro}', nome=_execucao)
        except:
            print('❌ Valor do FAP Original não encontrado')
            _escreve_relatorio_csv(f'{cnpj};Valor do FAP Original não encontrado', nome=_execucao)
            return 'não tem'

    return True


@_time_execution
def run():
    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--ignore-certificate-errors')
    # options.add_argument("--start-maximized")

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa

        _indice(count, total_empresas, empresa)
        
        login(cnpj, options)

    _escreve_header_csv('CNPJ;RESULTADO;DATA DO CÁLCULO;VIGÊNCIA', nome=_execucao)

if __name__ == '__main__':
    run()
