# -*- coding: utf-8 -*-
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
from pyautogui import prompt
from pathlib import Path
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start
from captcha_comum import _solve_text_captcha
from chrome_comum import _initialize_chrome


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


def localiza_tabela(empresa, driver, local_dados):
    # index é a quantidade referente ao tamanho da tabela, contem 12 linhas
    index = 12
    cnpj, senha, nome = empresa

    # para cada linha da tabela ele pega a informação de competência e o valor do caso houver
    # caso não exista nenhuma tabela e anotado o texto referente a situação do contribuinte
    # caso não consiga encontrar o texto sobre a situação do contribuinte retorna o erro
    for i in range(12):
        try:
            comp = driver.find_element(by=By.XPATH, value='//*[@id="j_idt13"]/table[2]/tbody/tr[{}]/td[1]/a'.format(index)).text
            resposta = driver.find_element(by=By.XPATH, value='//*[@id="j_idt13"]/table[2]/tbody/tr[{}]/td[3]'.format(index)).text
            try:
                quantidade = driver.find_element(by=By.XPATH, value='/html/body/form/table[2]/tbody/tr[{}]/td[2]'.format(index)).text
            except:
                quantidade = ''
            _escreve_relatorio_csv(';'.join([cnpj, nome, comp, resposta, quantidade]), local=local_dados)
            index -= 1
        except:
            try:
                soup = BeautifulSoup(driver.page_source, 'html.parser')
                padrao = re.compile(r'<tr><td class=\"coluna\"><span class=\"labelNegrito\">Situa.+ </span>(.+)</td><td class')
                resposta = padrao.search(str(soup))
                print(resposta.group(1))
                _escreve_relatorio_csv(';'.join([cnpj, nome, resposta.group(1)]), local=local_dados)
                break
            except:
                x = 'Erro'
                print('❌ Erro ao carregar as informações da empresa')
                return x

    x = 'Ok'
    print('✔ Consulta realizada')

    return x


def login(options, empresa, ano):
    #
    cnpj, senha, nome = empresa
    status, driver = _initialize_chrome(options)
    
    # entrar no site
    try:
        driver.get('http://gps.receita.fazenda.gov.br/')
        print('>>> Abrindo site...')
    except:
        print('❌ Time out\n>>> Tentando novamente...')
        x = 'Erro'
        return driver, x
    
    timer = 0
    while not find_by_id('captcha_challenge', driver):
        time.sleep(1)
        timer += 1
        if timer > 60:
            if not find_by_id('captcha_challenge', driver):
                print('❌ Timed out\n>>> Tentando novamente...')
                x = 'Erro'
                return driver, x
            
    # resolve o captcha
    captcha = _solve_text_captcha(driver, 'captcha_challenge')
    
    try:
        # insere o cnpj
        driver.find_element(by=By.ID, value='formInicio:identificador').send_keys(cnpj)
    
        # insere a competência final da consulta
        driver.find_element(by=By.ID, value='formInicio:competencia').send_keys('12{}'.format(ano))
    
        # insere a senha da empresa (geralmente e o começo do cnpj antes do 0001-XX)
        driver.find_element(by=By.ID, value='formInicio:senha').send_keys(senha)
    
        # insere a resposta do captcha
        driver.find_element(by=By.ID, value='captcha_campo_resposta').send_keys(captcha)
    
        time.sleep(0.5)
        # clica no botão para login
        driver.find_element(by=By.ID, value='formInicio:botaoConsultar').click()
        x = 'Ok'
    except:
        x = 'Erro'

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    padrao = re.compile(r'<h3 class=\"msg-erro\"><ul><li class=\"erro\">captcha:.+ (n.+).</li></ul>')
    resposta = padrao.search(str(soup))
    if resposta:
        print('❌ Erro na solução do captcha\n>>> Tentando novamente...')
        x = 'Erro'
        
    return driver, x


@_time_execution
def consulta():
    ano = prompt(text='Qual o ano da consulta', title='Script incrível', default='0000')
    empresas = _open_lista_dados()
    local_dados = Path(r'\\vpsrv03\Arq_Robo\Consulta Extrato INSS\ref. {}'.format(ano))

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # bloco para configurar o navegador, resolução padrão full hd e rodando sem abrir a tela do navegador
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")

    total_empresas = empresas[index:]

    # para cada empresa da lista de dados
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, senha, nome = empresa

        _indice(count, total_empresas, empresa, index)

        # faz 3 tentativas, se não der certo anota o erro na planilha
        x = 'Erro'
        tentativas = 1
        while x == 'Erro':
            # login na empresa
            driver, x = login(options, empresa, ano)

            # se o login retornar 'Ok'
            if x == 'Ok':
                index = 0
                print('>>> Aguardando site...')
                # espera a tabela aparecer
                while not find_by_id('j_idt13', driver):
                    time.sleep(1)
                    index += 1
                    if index > 60:
                        # se as tentativas excederem 3 sai do while e mantém o x = 'Ok' para sair do primeiro while
                        if tentativas >= 3:
                            _escreve_relatorio_csv(';'.join([cnpj, nome, 'Erro no login']), local=local_dados)
                            print('❌ Erro no login')
                            break
                        # se não conseguir e ainda estiver no limite de tentativas começa mais um siclo do while e retorna 'Erro
                        print('❌ Erro\n>>> Tentando novamente...')
                        x = 'Erro'
                        tentativas += 1
                        break

                # se achar a tabela percorre a tabela
                if find_by_id('j_idt13', driver):
                    x = localiza_tabela(empresa, driver, local_dados)
                    tentativas += 1

            driver.quit()
        index += 1


if __name__ == '__main__':
    consulta()
