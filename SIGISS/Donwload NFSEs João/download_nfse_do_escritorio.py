# -*- coding: utf-8 -*-
from requests import Session
from sys import path
import os, time
from selenium.webdriver.common.by import By

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice
from chrome_comum import initialize_chrome


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None
    
    
def salvar_arquivo(nome, response):
    try:
        arquivo = open(os.path.join('execucao/Documentos', nome), 'wb')
    except FileNotFoundError:
        os.makedirs('execucao/Documentos', exist_ok=True)
        arquivo = open(os.path.join('execucao/Documentos', nome), 'wb')
    for parte in response.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()


def download(empresas, index):
    query = {'loginacesso': '24982859000101',
             'senha': '2376500'}
    url = 'https://valinhos.sigissweb.com/'

    status, driver = initialize_chrome()

    driver.get(url)
    
    while not find_by_id('edtlogin', driver):
        time.sleep(1)
        
    element = driver.find_element(by=By.ID, value='edtlogin')
    element.send_keys('24982859000101')
    
    element = driver.find_element(by=By.ID, value='edtsenha')
    element.send_keys('2376500')
    
    time.sleep(2)
    
    button = driver.find_element(by=By.XPATH, value='//*[@id="acesso"]/fieldset/input[3]')
    button.click()
    
    while not find_by_id('nav', driver):
        time.sleep(1)
        
    driver.get('https://valinhos.sigissweb.com/nfecentral?oper=pesquisanfe')
    
    while not find_by_id('data1', driver):
        time.sleep(1)
    
    time.sleep(2)
    element = driver.find_element(by=By.ID, value='data1')
    element.send_keys('01/07/2022')

    element = driver.find_element(by=By.XPATH, value='//*[@id="operData"]')
    element.click()
    element = driver.find_element(by=By.XPATH, value='//*[@id="operData"]')
    element.send_keys('Menor que')
    element = driver.find_element(by=By.XPATH, value='//*[@id="operData"]')
    element.click()

    element = driver.find_element(by=By.ID, value='btnexecutar')
    element.click()
    
    time.sleep(5)
    
    element = driver.find_element(by=By.XPATH, value='//a[@href=' + "javascript: abrePagina( 'cadnfegrid', 'imprimirnfe', '&cnpj=24982859000101&numeronf=8001&serienf=NFD&tipo=I' )" + ']')
    element.click()
    
    time.sleep(40)
    total_empresas = empresas[index:]
# https://valinhos.sigissweb.com/nfecentral?oper=imprimir&cnpj=24982859000101&numeronf=8001&serienf=FND&tipo=l
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        num, chave = empresa
        
        with Session() as s:
            # entra no site
            s.get('https://valinhos.sigissweb.com/')

            # loga na empresa
            query = {'loginacesso': cnpj,
                     'senha': senha}
            s.post('https://valinhos.sigissweb.com/ControleDeAcesso', data=query)
        
        js = "javascript: abrePagina( 'cadnfegrid', 'imprimirnfe', '&cnpj=24982859000101&numeronf=8001&serienf=NFD&tipo=I' )"
        
        response = driver.execute_script(js)
        print(response)
        '''# Salva o PDF da nota e anota na planilha
        salvar_arquivo('nfe_' + num + '.pdf', response)
        _escreve_relatorio_csv(';'.join([num, 'Nota Fiscal salva']), 'Download NFSEs do Escritório.csv')
        print('✔ Download concluído, nota: ' + num)
    
        # Se der erro para validar a nota
        print('❌ ERRO, nota: ' + num)
        _escreve_relatorio_csv(';'.join([num, 'Erro']), 'Download NFSEs do Escritório.csv')
        continue'''


@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    download(empresas, index)


if __name__ == '__main__':
    run()
