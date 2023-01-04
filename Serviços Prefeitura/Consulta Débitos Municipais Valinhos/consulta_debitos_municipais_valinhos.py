# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
import os, time, re
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha

os.makedirs('execução/documentos', exist_ok=True)


def reset_table(table):
    table.attrs = None
    for tag in table.findAll(True):
        tag.attrs = None
        try:
            tag.string = tag.string.strip()
        except:
            pass

    return table


def format_data(content):
    # tabelas = []
    soup = BeautifulSoup(content, 'html.parser')
    tabela_header = soup.find('table', attrs={'id': 'table6'})
    tabela_debitos = soup.find('div', attrs={'id': 'grid'}).find('table')
    tabela_totais = soup.find('table', attrs={'id': 'tableTotais'})
    
    linhas = tabela_header.findAll('td', attrs={'class': 'linhaInformacao'})
    paragrafo = ', '.join(linha.text for linha in linhas)

    if len(tabela_debitos.findAll('tr')) <= 1:
        return False
    # tabela_debitos.attrs = None
    for tag in tabela_debitos.findAll(True):
        # tag.attrs = None
        try:
            tag.string = tag.string.strip()
        except:
            pass
        if 'input' == tag.name:
            tag.parent.extract()

    tabela_totais.find('tr', attrs={'id': 'tr4'}).extract()
    tabela_totais.find('td', attrs={'id': 'td2'}).extract()
    tabela_totais.find('td', attrs={'id': 'td4'}).extract()
    # tabela_totais = reset_table(tabela_totais)

    with open('style.css', 'r') as f:
        css = f.read()

    style = "<style type='text/css'>" + css + "</style>"
    new_soup = BeautifulSoup(f'<html><head><meta charset="UTF-8">{style}</head><body><p></p></body></html>', 'html.parser')
    new_soup.body.p.append(paragrafo)
    new_soup.body.append(tabela_debitos)
    new_soup.body.append(tabela_totais)

    return str(new_soup)


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def login(options, cnpj, insc_muni):
    status, driver = _initialize_chrome(options)
    base = 'http://179.108.81.10:9081/tbw'
    url_inicio = f'{base}/loginWeb.jsp?execobj=ServicoPesquisaISSQN'

    driver.get(url_inicio)
    while not find_by_id('span7Menu', driver):
        time.sleep(1)
    button = driver.find_element(by=By.ID, value='span7Menu')
    button.click()

    while not find_by_id('captchaimg', driver):
        time.sleep(1)
    element = driver.find_element(by=By.ID, value='captchaimg')
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
    captcha = _solve_text_captcha(os.path.join('ignore', 'captcha', 'captcha.png'))

    if not captcha:
        print('Erro Login - não encontrou captcha')
        return 'erro captcha'

    if captcha is not None:
        while not find_by_id('input1', driver):
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input1')
        element.send_keys(insc_muni)

        while not find_by_id('input4', driver):
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input4')
        element.send_keys(cnpj)

        while not find_by_id('captchafield', driver):
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='captchafield')
        element.send_keys(captcha)

        while not find_by_id('imagebutton1', driver):
            time.sleep(1)
        button = driver.find_element(by=By.ID, value='imagebutton1')
        button.click()

        time.sleep(3)
        
        print('>>> Aguardando site')
        timer = 0
        while not find_by_id('td30', driver):
            driver.save_screenshot(r'ignore\debug_screen.png')
            time.sleep(1)

            if find_by_id('tdMsg', driver):
                try:
                    erro = re.compile(r'\"erroMessageText\".+tipo=\"td\">(.+)</td></tr></tbody></table>').search(driver.page_source).group(1)
                    if erro != 'td':
                        _escreve_relatorio_csv(f'{cnpj};{erro}')
                        print(f'❌ {erro}')
                        return False
                except:
                    pass
            timer += 1
            if timer >= 10:
                _escreve_relatorio_csv(f'{cnpj};Erro no login')
                print(f'❌ Erro no login')
                return False
            
        html = driver.page_source.encode('utf-8')
        try:
            driver.save_screenshot(r'ignore\debug_screen.png')
            str_html = format_data(html)

            if str_html:
                _escreve_relatorio_csv(';'.join([cnpj, 'Com débitos']))
                print('>>> Gerando arquivo')
                nome_arq = ';'.join([cnpj, 'INF_FISC_REAL', 'Debitos Municipais'])
                with open('execução/documentos/' + nome_arq + r'.pdf', 'w+b') as pdf:
                    pisa.showLogging()
                    pisa.CreatePDF(str_html, pdf)
                    print('❗ Arquivo gerado')
            else:
                _escreve_relatorio_csv(';'.join([cnpj, 'Sem débitos']))
                print('✔ Não há débitos')
        except:
            _escreve_relatorio_csv(';'.join([cnpj, 'Erro na geração do PDF']))
            driver.save_screenshot(r'ignore\debug_screen.png')
            print('❌ Erro na geração do PDF')

        driver.close()
        return True
    else:
        _escreve_relatorio_csv(f'{cnpj};Erro no login')
        print('❌ Erro no login')
        return True


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
        cnpj, insc_muni = empresa

        _indice(count, total_empresas, empresa)

        resultado = 'erro captcha'
        while resultado == 'erro captcha':
            resultado = login(options, cnpj, insc_muni)


if __name__ == '__main__':
    run()
