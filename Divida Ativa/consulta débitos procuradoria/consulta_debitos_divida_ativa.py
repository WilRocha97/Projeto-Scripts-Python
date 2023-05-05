# -*- coding: utf-8 -*-
import time
import re
from bs4 import BeautifulSoup
from requests import Session
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_recaptcha
from chrome_comum import _initialize_chrome
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
import os


def verifica_debitos(pagina):
    try:
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # debito = soup.find('span', attrs={'id': 'consultaDebitoForm:dataTable:tb'})
        soup = soup.prettify()
        request_verification = re.compile(r'(consultaDebitoForm:dataTable:tb)').search(soup).group(1)
        return request_verification
    except:
        re.compile(r'(Nenhum resultado com os critérios de consulta)').search(soup).group(1)
        return False


def limpa_registros(html):
    linhas = html.findAll('tr')
    for i in linhas:
        colunas = i.findAll('td')
        if len(colunas) >= 7:
            # Remove informações de pagamentos
            if colunas[5].text.startswith('Liquidar'):
                for a in colunas[5].findAll('a'):
                    a.extract()
            # Remove informações de emissão de GARE
            if colunas[6].find('table'):
                for t in colunas[6].findAll('table'):
                    t.extract()
    
    return html


def salva_pagina(pagina, cnpj, compl=''):
    os.makedirs('execução\documentos', exist_ok=True)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    formulario = soup.find('form', attrs={'id': 'consultaDebitoForm'})
    
    imagens = formulario.findAll('img')
    for img in imagens:
        img.extract()
    botoes = formulario.findAll('input', attrs={'type': 'submit'})
    for botao in botoes:
        botao.extract()
    
    with open('style.css', 'r') as f:
        css = f.read()
    style = "<style type='text/css'>" + css + "</style>"
    html = BeautifulSoup(f'<html><head><meta charset="UTF-8">{style}</head><body></body></html>', 'html.parser')
    
    # cabeçalho
    try:
        parte = formulario.find('div', attrs={'id': 'consultaDebitoForm:j_id127_header'})
        html.body.append(parte)
    except:
        pass
    # devedor
    try:
        parte = formulario.find('span', attrs={'id': 'consultaDebitoForm:consultaDevedor'}).find('table')
        html.body.append(parte)
    except:
        pass
    # titulo tabela
    try:
        parte = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTableDebitosRelativo'})
        html.body.append(parte)
    except:
        pass
    # tabela principal
    try:
        new_tag = BeautifulSoup('<table class="main-table"></table>', 'html.parser')
        head = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('thead')
        body = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('tbody')
        body = limpa_registros(body)
        footer = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('tfoot')
        new_tag.table.append(head)
        new_tag.table.append(body)
        new_tag.table.append(footer)
        html.body.append(new_tag)
    except:
        pass
    
    debito = f' - {compl}' if compl else ''
    nome = os.path.join('execução\documentos', f'{cnpj};INF_FISC_REAL;Procuradoria Federal{debito}.pdf')
    with open(nome, 'w+b') as pdf:
        # pisa.showLogging()
        pisa.CreatePDF(str(html), pdf)
    print('✔ Arquivo salvo')
    _escreve_relatorio_csv(f'{cnpj};Empresa com debitos;{compl}')
    
    return True


def inicia_sessao():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = _initialize_chrome(options)
    
    url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
    driver.get(url)
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
    token = _solve_recaptcha(recaptcha_data)
    
    driver.find_element(by=By.ID, value='consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa').click()
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div/div[2]/div/div[3]/div/div[2]/div[2]/span/form/span/div/div[2]/table/tbody/tr/td[1]/div/div/span/select/option[2]').click()
    time.sleep(1)
    
    driver.find_element(by=By.ID, value='consultaDebitoForm:decTxtTipoConsulta:cnpj').send_keys('00656064000145')
    driver.execute_script("document.getElementById('g-recaptcha-response').innerText='" + token + "'")
    time.sleep(1)
    
    driver.execute_script("document.getElementsByName('consultaDebitoForm:j_id102')[0].click()")
    time.sleep(1)
    return driver


def consulta_debito(s, empresa):
    cnpj, nome = empresa
    url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
    # str_cnpj = f"{cnpj[0:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    pagina = s.get(url)
    
    try:
        soup = BeautifulSoup(pagina.content, 'html.parser')
        viewstate = soup.find(id="javax.faces.ViewState").get('value')
    except Exception as e:
        print('❌ Não encontrou viewState')
        print(e)
        s.close()
        return False
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
    token = _solve_recaptcha(recaptcha_data)
    
    # Troca opção de pesquisa para CNPJ
    info = {
        'AJAXREQUEST': '_viewRoot',
        'consultaDebitoForm': 'consultaDebitoForm',
        'consultaDebitoForm:decTxtTipoConsulta:cdaEtiqueta': '',
        'g-recaptcha-response': '',
        'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
        'consultaDebitoForm:modalDadosCartorioOpenedState': '',
        'javax.faces.ViewState': viewstate,
        'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
        'ajaxSingle': 'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa',
        'consultaDebitoForm:decLblTipoConsulta:j_id76': 'consultaDebitoForm:decLblTipoConsulta:j_id76'
    }
    s.post(url, info)
    
    # Consulta o cnpj
    info = {
        'consultaDebitoForm': 'consultaDebitoForm',
        'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
        'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
        'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
        'g-recaptcha-response': token,
        'consultaDebitoForm:j_id104': 'Consultar',
        'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
        'consultaDebitoForm:modalDadosCartorioOpenedState': '',
        'javax.faces.ViewState': viewstate
    }
    pagina = s.post(url, info)
    debitos = verifica_debitos(pagina)
    
    if not debitos:
        print('✔ Sem débitos')
        _escreve_relatorio_csv(f'{cnpj};Empresa sem debitos')
    else:
        print('❗ Com débitos')
        soup = BeautifulSoup(pagina.content, 'html.parser')
        tabela = soup.find('tbody', attrs={'id': 'consultaDebitoForm:dataTable:tb'})
        linhas = tabela.find_all('a')
        viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
        for index, linha in enumerate(linhas):
            tipo = linha.get('id')
            info = {
                'consultaDebitoForm': 'consultaDebitoForm',
                'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
                'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
                'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
                'g-recaptcha-response': token,
                'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
                'consultaDebitoForm:modalDadosCartorioOpenedState': '',
                'javax.faces.ViewState': viewstate,
                f'consultaDebitoForm:dataTable:{index}:lnkConsultaDebito': tipo
            }
            pagina = s.post(url, info)
            salva_pagina(pagina, cnpj, linha.text)
            
            # Retorna para tela de consulta
            viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
            info = {
                'consultaDebitoForm': 'consultaDebitoForm',
                'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
                'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
                'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
                'g-recaptcha-response': token,
                'consultaDebitoForm:j_id260': 'Voltar',
                'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
                'consultaDebitoForm:modalDadosCartorioOpenedState': '',
                'javax.faces.ViewState': viewstate
            }
            s.post(url, info)
            viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
    
    return True


@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    driver = inicia_sessao()

    cookies = driver.get_cookies()
    driver.quit()
    with Session() as s:
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])
            
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            _indice(count, total_empresas, empresa)
            
            erro = 'sim'
            while erro == 'sim':
                try:
                    consulta_debito(s, empresa)
                    erro = 'nao'
                except:
                    erro = 'sim'
    s.close()
    

if __name__ == '__main__':
    run()
