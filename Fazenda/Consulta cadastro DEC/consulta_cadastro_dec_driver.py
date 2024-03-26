import time, re, os, pyautogui as p
from xhtml2pdf import pisa
from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from captcha_comum import _solve_hcaptcha
from fazenda_comum import _atualiza_info
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def login(driver, cnpj_formatado):
    while True:
        url = 'https://www.dec.fazenda.sp.gov.br/DEC/UCConsultaPublica/Consulta.aspx'
        recaptcha_data = {'sitekey': '95ac31b4-a1cb-4e67-8405-aac0bca555da', 'url': url}
        
        driver.get(url)
        id_recaptcha = re.compile(r'id=\"(g-recaptcha-response-.+)\" name=\"g-recaptcha-response\"').search(driver.page_source).group(1)
        id_hcaptcha = re.compile(r'id=\"(h-captcha-response-.+)\" name=\"h-captcha-response\"').search(driver.page_source).group(1)
        
        captcha = _solve_hcaptcha(recaptcha_data, visible=True)
        
        # inserir captcha
        driver.execute_script("document.getElementById('{}').innerHTML='{}';".format(id_recaptcha, captcha))
        driver.execute_script("document.getElementById('{}').innerHTML='{}';".format(id_hcaptcha, captcha))
        driver.find_element(by=By.ID, value='ConteudoPagina_txtEstabelecimentoBusca').send_keys(cnpj_formatado)
        
        
        
        driver.find_element(by=By.ID, value='ConteudoPagina_btnBuscarPorEstabelecimento').click()
        
        print(driver.page_source)
        time.sleep(44)


@_time_execution
def run():
    andamentos = 'Consulta Situação DEC'
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome= empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        status, driver = _initialize_chrome(options)
        
        cnpj_formatado = re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', cnpj)
        
        resultado = login(driver, cnpj_formatado)
        
    
    
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;NOME;SEMESTRE;ANO;SALDO', nome=andamentos)


if __name__ == '__main__':
    run()
