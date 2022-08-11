# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from requests import Session
from PIL import Image
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import initialize_chrome
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha

os.makedirs('execucao/documentos', exist_ok=True)


def login(cnpj, insc_muni):

    print('>>> Abrindo o site')
    url = 'https://jundiai.sp.gov.br/servicos-online/certidao-negativa-de-debitos-mobiliarios/'
    url2 = 'https://web.jundiai.sp.gov.br/PMJ/SW/certidaonegativamobiliario.aspx'

    with Session() as s:
        # s.get(url)
        s.get(url)
        pagina = s.get(url2)
        
        soup = BeautifulSoup(pagina.content, 'html.parser')
        viewstate = soup.find(attrs={'id': "__VIEWSTATE"})['value']
        viewstate_generator = soup.find(attrs={'id': "__VIEWSTATEGENERATOR"})['value']
        event_validation = soup.find(attrs={'id': "__EVENTVALIDATION"})['value']

        recaptcha_data = {'sitekey': '6LfK0igTAAAAAOeUqc7uHQpW4XS3EqxwOUCHaSSi', 'url': url2}
        token = _solve_recaptcha(recaptcha_data)
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36'
        }
        
        info = {
            'ScriptManager1': 'UpdatePanel1 | btnEnviar',
            '__LASTFOCUS': '',
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            '__VIEWSTATE': viewstate,
            '__VIEWSTATEGENERATOR': viewstate_generator,
            '__EVENTVALIDATION': event_validation,
            'DadoContribuinteMobiliario1$txtCfm': insc_muni,
            'DadoContribuinteMobiliario1$txtNrCicCgc': cnpj,
            'g-recaptcha - response': token,
            '__ASYNCPOST': 'true',
            'btnEnviar': 'Consultar'
        }
        
        print('>>> Logando na empresa')
        pagina = s.post('https://web.jundiai.sp.gov.br/PMJ/SW/certidaonegativamobiliario.aspx', data=info, headers=headers)
        # pagina = s.get('https://web.jundiai.sp.gov.br/PMJ/SW/Solicitante.aspx?value=Q6C1oKOlOI3FsNGfqTs8gOmvEI8l7bc6RlQl27MkHS%2fKElZq33QUlyv2pkK%2bMyr4gJnKOnX4xHShmXZKJnM3GtXMXiUyrubP')
        soup = BeautifulSoup(pagina.content, 'html.parser')
        soup = soup.prettify()

        print('>>> Consultando empresa')
        print(soup)
        situacao = re.compile(r'<span id=\"lblSituacao\">\n +(.+)')
        situacao = situacao.search(soup)
        situacao = situacao.group(1)
        
        situacao_print = f'❌ {situacao}'
    
        return situacao, situacao_print


@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, insc_muni = empresa

        _indice(count, total_empresas, empresa)

        situacao, situacao_print = login(cnpj, insc_muni)
        _escreve_relatorio_csv(f'{cnpj};{insc_muni};{situacao}')
        print(situacao_print)


if __name__ == '__main__':
    run()
