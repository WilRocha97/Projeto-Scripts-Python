# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from sys import path
from selenium import webdriver
import time, re

path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _time_execution, _open_lista_dados, _where_to_start
from chrome_comum import _initialize_chrome

site = 'https://valoresareceber.bcb.gov.br/publico/'


def find_by_id(html_id, driver):
    try:
        elem = driver.find_element_by_id(html_id)
        return elem
    except:
        return None


def consultar(empresas, index):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        info, data = empresa

        _indice(count, total_empresas, empresa)

        # não usa as custom options porque o site não carrega com elas
        status, driver = _initialize_chrome()

        driver.get(site)
        print('>>> Entrando no site')

        i = 0
        while not find_by_id('cpf', driver):
            if i >= 10:
                driver.get(site)
                i = 0
            time.sleep(1)
            i += 1

        if len(info) > 11:
            button = driver.find_element_by_xpath('//*[@id="top"]/div/app-consulta-publica/div/app-consulta/div/div/div/form/div[2]/div/div/div[1]/div[2]/label')
            button.click()

        element = driver.find_element_by_xpath('/html/body/app-root/div/app-main/main/div/app-consulta-publica/div/app-consulta/div/div/div/form/div[2]/div/div/div[2]/div[1]/input')
        element.send_keys(info)

        element = driver.find_element_by_id('dataNascimento')
        element.send_keys(data)

        button = driver.find_element_by_xpath('//*[@id="top"]/div/app-consulta-publica/div/app-consulta/div/div/div/form/div[3]/button')
        button.click()
        print('>>> Realizando a consulta')
        time.sleep(3)

        html = driver.page_source.encode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')

        # padrao_negativo = re.compile(r'id=\"consultaNegativa\".+> (.+) </h1>')
        padrao_positivo = re.compile(r'Favor voltar neste mesmo site.+</p><div><ul><li> (.+) </li></ul>')
        padrao_repescagem = re.compile(r'Repescagem:.+</div><div><li> (.+) </li>')
        padrao_erro = re.compile(r'Documento informado é inválido')

        evento_erro = padrao_erro.search(str(soup))

        if evento_erro:
            resultado = str(evento_erro.group())
            print('❌ {}'.format(resultado))
            texto = '{};{};{}'.format(info, data, resultado)
            _escreve_relatorio_csv(texto)
            driver.close()
            continue

        evento_sn = padrão_negativo.search(str(soup))
        if evento_sn:
            resultado = str(evento_sn.group(1))
            print('❌ {}'.format(resultado))

        else:
            evento_sn = padrao_positivo.search(str(soup))
            evento_repescagem = padrao_repescagem.search(str(soup))
            resultado = 'Solicitar o resgate em {};Repescagem em {}'.format(str(evento_sn.group(1)), str(evento_repescagem.group(1)))
            print('✔ {}'.format(resultado))

        texto = '{};{};{}'.format(info, data, resultado)
        _escreve_relatorio_csv(texto)

        driver.close()


@_time_execution
def run():
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
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    consultar(empresas, index)


if __name__ == '__main__':
    run()
