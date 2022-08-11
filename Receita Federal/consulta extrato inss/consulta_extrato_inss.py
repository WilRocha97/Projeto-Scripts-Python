# -*- coding: utf-8 -*-
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from PIL import Image
from comum import initialize_chrome
from sys import path
import os, time, re

path.append(r'..\..\_comum')
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start
from captcha_comum import _break_normal_captcha


def find_by_id(info_id, driver):
    # função para o find_by_id não parar o robô
    try:
        elem = driver.find_element_by_id(info_id)
        return elem
    except NoSuchElementException:
        return None


def find_by_xpath(info_path, driver):
    # função para o find_by_xpath não parar o robô
    try:
        elem = driver.find_element_by_id(info_path)
        return elem
    except NoSuchElementException:
        return None


def localiza_tabela(empresa, driver):
    # index é a quantidade referênte ao tamanho da tabela, contem 12 linhas
    index = 12
    cnpj, senha, nome = empresa

    # para cada linha da tabela ele pega a informação de competência e o valor do caso houver
    # caso não exista nenhuma tabela e anotado o texto referênte a situação do contribuinte
    # caso não consiga encontrar o texto sobre a situação do contribuite retorna o erro
    time.sleep(60)
    for i in range(12):
        try:
            comp = driver.find_element_by_xpath('//*[@id="j_idt13"]/table[2]/tbody/tr[{}]/td[1]/a'.format(index)).text
            resposta = driver.find_element_by_xpath('//*[@id="j_idt13"]/table[2]/tbody/tr[{}]/td[3]'.format(index)).text
            _escreve_relatorio_csv(';'.join([cnpj, nome, comp, resposta]))
            index -= 1
        except:
            html = driver.page_source
            try:
                soup = BeautifulSoup(html, 'html.parser')
                padrao = re.compile(r'<tr><td class=\"coluna\"><span class=\"labelNegrito\">Situa.+ </span>(.+)</td><td class')
                resposta = padrao.search(str(soup))
                print(resposta.group(1))
                _escreve_relatorio_csv(';'.join([cnpj, nome, resposta.group(1)]))
                break
            except:
                x = 'Erro'
                print('❌ Erro ao carregar as informações da empresa')
                return x

    x = 'Ok'
    print('✔ Consulta realizada')

    return x


def login(driver, empresa, ano):
    #
    cnpj, senha, nome = empresa

    # entrar no site
    try:
        driver.get('http://gps.receita.fazenda.gov.br/')
        print('>>> Abrindo site...')
    except:
        print('❌ Time out\n>>> Tentando novamente...')
        x = 'Erro'
        return driver, x

    # enquanto site não abre, espere 120 segundos e se depois disso não aparecer o captcha para resolver, tenta novamente
    while not find_by_id('captcha_challenge', driver):
        time.sleep(120)
        if not find_by_id('captcha_challenge', driver):
            driver.delete_all_cookies()
            print('>>> Recarregando a página...')
            try:
                driver.get('http://gps.receita.fazenda.gov.br/')
            except:
                print('❌ Timed out\n>>> Tentando novamente...')
                x = 'Erro'
                return driver, x

    # pega a imagem do captcha
    os.makedirs(r'captcha', exist_ok=True)
    name = os.path.join(r'captcha', 'captcha.png')

    x, y, width, height = driver.find_element_by_id('captcha_challenge').rect.values()
    x, y, width, height = width, height, width + y, height + x

    driver.save_screenshot(name)
    with Image.open(name) as img:
        aux = img.crop((x, y, width, height))
        aux.save(name)

    # resolve o captcha
    captcha = _break_normal_captcha(os.path.join('captcha', 'captcha.png'), limit=1)

    # se o captcha for resolvido
    if captcha is not str:
        # insere o cnpj
        try:
            driver.find_element_by_id('formInicio:identificador').send_keys(cnpj)

            # insere a competência final da consulta
            driver.find_element_by_id('formInicio:competencia').send_keys('12{}'.format(ano))

            # insere a senha da empresa (geralmente e o começo do cnpj antes do 0001-XX)
            driver.find_element_by_id('formInicio:senha').send_keys(senha)

            # insere a resposta do captcha
            driver.find_element_by_id('captcha_campo_resposta').send_keys(captcha['code'])

            time.sleep(0.5)
            # clica no botão para login
            driver.find_element_by_id('formInicio:botaoConsultar').click()
            x = 'Ok'
        except:
            x = 'Erro'
            return driver, x

        html = driver.page_source.encode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        padrao = re.compile(r'<h3 class=\"msg-erro\"><ul><li class=\"erro\">captcha:.+ (n.+).</li></ul>')
        resposta = padrao.search(str(soup))
        if resposta:
            print('❌ Erro na solução do captcha\n>>> Tentando novamente...')
            x = 'Erro'
            return driver, x

        return driver, x

    # se o captcha não for resolvido
    else:
        print('❌ Erro na solução do captcha\n>>> Tentando novamente...')
        x = 'Erro'
        return driver, x


@_time_execution
def consulta():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # bloco para configurar o navegador, resolução padrão full hd e rodando sem abrir a tela do navegador
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    status, driver = initialize_chrome()

    total_empresas = empresas[index:]
    ano = datetime.now().year
    # para cada empresa da lista de dados
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, senha, nome = empresa

        _indice(count, total_empresas, empresa)

        # faz 3 tentativas, se não der certo anota o erro na planilha
        x = 'Erro'
        tentativas = 1
        while x == 'Erro':
            # login na empresa
            driver, x = login(driver, empresa, ano)

            # se o login retornar 'Ok'
            if x == 'Ok':
                index = 1
                print('>>> Aguardando site...')
                # espera a tabela aparecer
                while not find_by_id('j_idt13', driver):
                    time.sleep(1)
                    index += 1
                    if index > 60:
                        # se as tentativas excederem 3 sai do while e mantem o x = 'Ok' para sair do primeiro while
                        if tentativas >= 3:
                            _escreve_relatorio_csv(';'.join([cnpj, nome, 'Erro no login']))
                            print('❌ Erro no login')
                            break
                        # se não conseguir e ainda estiver dentro do limite de tentativas começa mais um siclo do while e retorna 'Erro
                        print('❌ Erro\n>>> Tentando novamente...')
                        x = 'Erro'
                        tentativas += 1
                        break

                # se achar a tabela percorre a tabela
                if find_by_id('j_idt13', driver):
                    x = localiza_tabela(empresa, driver)
                    tentativas += 1

            # deleta os cookies e cache do site
            driver.delete_all_cookies()
        index += 1


if __name__ == '__main__':
    consulta()
