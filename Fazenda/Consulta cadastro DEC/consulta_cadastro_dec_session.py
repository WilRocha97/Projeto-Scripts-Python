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


def login(cnpj_formatado):
    while True:
        
        url = 'https://www.dec.fazenda.sp.gov.br/DEC/UCConsultaPublica/Consulta.aspx'
        recaptcha_data = {'sitekey': '95ac31b4-a1cb-4e67-8405-aac0bca555da', 'url': url}
        captcha = _solve_hcaptcha(recaptcha_data)
        with Session() as s:
            # entra no site
            pagina = s.get('https://www.dec.fazenda.sp.gov.br/DEC/UCConsultaPublica/Consulta.aspx', verify=False)
            
            query = {'response': captcha}
            s.post('https://www.dec.fazenda.sp.gov.br/DEC/UCConsultaPublica/Consulta.aspx/VerifyCaptcha', data=query)
            
            state, generator, validation = _atualiza_info(pagina)
            
            print('>>> Logando no usuário')
            # loga no usuario
            query = {'__LASTFOCUS': '',
                    '__EVENTTARGET': '',
                    '__EVENTARGUMENT': '',
                    '__VIEWSTATE': state,
                    '__VIEWSTATEGENERATOR': generator,
                    '__EVENTVALIDATION': validation,
                    'menu1': 'Destaques',
                    'ctl00$ConteudoPagina$txtEstabelecimentoBusca': cnpj_formatado,
                    'g-recaptcha-response': captcha,
                    'h-captcha-response': captcha,
                    'ctl00$ConteudoPagina$txtCaptcha': '',
                    'ctl00$ConteudoPagina$btnBuscarPorEstabelecimento': 'Buscar'}
            
            res = s.post('https://www.dec.fazenda.sp.gov.br/DEC/UCConsultaPublica/Consulta.aspx', data=query)
            
            soup = BeautifulSoup(res.content, 'html.parser')
            soup = soup.prettify()
            
            
            
            if re.compile(r'É necessário realizar a validação do hCaptcha').search(soup):
                print('>>> Falha na resolução do Captcha, tentando novamente.')
                s.close()
            else:
                print(soup)
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
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    s = False
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome= empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        cnpj_formatado = re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', cnpj)
        
        resultado = login(cnpj_formatado)
        
    
    
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;NOME;SEMESTRE;ANO;SALDO', nome=andamentos)


if __name__ == '__main__':
    run()
