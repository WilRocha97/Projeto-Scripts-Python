# -*- coding: utf-8 -*-
import time, re, os
from selenium import webdriver
from PIL import Image
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha


def login(driver, nome, cpf, pis, data_nasc):
    print('>>> Acessando site')
    try:
        driver.get('http://consultacadastral.inss.gov.br/Esocial/pages/index.xhtml')
    except:
        return driver, 'erro'
    
    while not _find_by_id('indexForm1:botaoConsultar', driver):
        driver.get('http://consultacadastral.inss.gov.br/Esocial/pages/index.xhtml')
        time.sleep(1)
        
    driver.find_element(by=By.ID, value='indexForm1:botaoConsultar').click()
    
    while not _find_by_id('formQualificacaoCadastral:nome', driver):
        time.sleep(1)
    
    itens = [('formQualificacaoCadastral:nome', nome),
             ('formQualificacaoCadastral:dataNascimento', data_nasc),
             ('formQualificacaoCadastral:cpf', cpf),
             ('formQualificacaoCadastral:nis', pis)]
    
    for iten in itens:
        driver.find_element(by=By.ID, value=iten[0]).click()
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])
    
    print('>>> Acessando cadastro')
    driver.find_element(by=By.ID, value='formQualificacaoCadastral:btAdicionar').click()
    
    while not _find_by_id('gridDadosTrabalhador', driver):
        time.sleep(1)
    
    driver.find_element(by=By.ID, value='formValidacao2:botaoValidar2').click()
    
    # tira um print da tela com e recorta apenas a imagem do captcha para enviar para a api
    while not _find_by_id('captcha_challenge', driver):
        time.sleep(1)
    element = driver.find_element(by=By.ID, value='captcha_challenge')
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
        return driver, 'erro captcha'
    
    driver.find_element(by=By.ID, value='captcha_campo_resposta').send_keys(captcha)
    driver.find_element(by=By.ID, value='formValidacao:botaoValidar').click()
    
    print('>>> Acessando cadastro')
    timer = 0
    while not _find_by_id('j_idt21:gridDadosTrabalhador', driver):
        time.sleep(1)
        timer += 1
        if re.compile(r'captcha: A resposta .+ não corresponde ao desafio gerado').search(driver.page_source):
            print('❌ A resposta não corresponde ao desafio gerado, tentando novamente')
            return driver, 'erro'
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente')
            return driver, 'erro'
        
    return driver, 'ok'


def consulta(driver):
    try:
        mensagem = re.compile(r'</span><span class=\"tamanho\d+.+>(.+)<br><br></span></td><td').search(driver.page_source).group(1)
        return driver, str(mensagem)
    except:
        pass
    
    mensagens_regex = [r'</span><span class=\"tamanho\d+.+>(.+)<br><br></span></td><td class=\"left\"><span class=\"tamanho\d+\"></span><span class=\"tamanho\d+\">(.+)<br><br> </span></td></tr>',
                       r'<span class=\"tamanho\d+.+>(.+)<br><br></span><span class=\"tamanho\d+\"></span></td><td class=\"left\"><span class=\"tamanho\d+\">(.+)<br><br></span><span class=\"tamanho\d+\"> </span></td></tr>']
    
    for mensagem_regex in mensagens_regex:
        try:
            mensagens = re.compile(mensagem_regex).search(driver.page_source)
            mensagem = f'{mensagens.group(1)};{mensagens.group(2).replace(":<br>", " | ").replace("<br>", " | ")}'
            return driver, str(mensagem)
        except:
            pass
    
    print(driver.page_source)
    return driver, str('Erro ao analisar o cadastro')

@_time_execution
def run():
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1366,768')
    # options.add_argument("--start-maximized")
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        cpf, nome, cod_empresa, cod_empregado, pis, data_nasc = empresa
        
        while 1 > 0:
            # iniciar o driver do chome
            status, driver = _initialize_chrome(options)
            
            driver, resultado = login(driver, nome, cpf, pis, data_nasc)
            if resultado != 'erro':
                break
            driver.close()
                
        driver, resultado = consulta(driver)
        print(f'❕ {resultado}')
        _escreve_relatorio_csv(f'{cpf};{nome};{cod_empresa};{cod_empregado};{pis};{data_nasc};{resultado}')
        driver.close()

if __name__ == '__main__':
    run()