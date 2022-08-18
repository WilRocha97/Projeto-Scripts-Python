# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from sys import path
from selenium import webdriver
from urllib3 import disable_warnings, exceptions

path.append(r'..\..\_comum')
from sn_comum import _new_session_sn
from pyautogui_comum import _get_comp
from chrome_comum import _initialize_chrome
from comum_comum import _time_execution, _download_file, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice


def consulta_recibos(cnpj, mes, ano, session):
    url_base = 'https://www8.receita.fazenda.gov.br'
    url_cons = f'{url_base}/SimplesNacional/Aplicacoes/ATSPO/pgdasd2018.app/Consulta'

    disable_warnings(exceptions.InsecureRequestWarning)
    
    session.get(url_cons, verify=False)
    res = session.post(url_cons, {'anoDigitado': ano}, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')
    # procura no código do site se existe recibo
    recibos = soup.findAll('a', attrs={'data-content': 'Imprimir Recibo'})

    download = (0, 0)
    for recibo in recibos:
        link = recibo.get('href', '')
        if not link:
            continue

        # pega a id da declaração que possuí o recibo
        doc = link.split('idDeclaracao=')[1][:17]
        doc_ano, doc_mes, doc_num = doc[8:12], doc[12:14], int(doc[14:17])

        if f'{mes}/{ano}' != f'{doc_mes}/{doc_ano}':
            continue
        if doc_num < download[0]:
            continue
        
        # faz o download do recibo
        download = (doc_num, link)

    # se o número do documento for igual a 0 retorna que o o recibo não foi encontrado
    if download[0] == 0:
        print(f'❗ Não encontrou recibo {mes}/{ano}')
        return f'Não encontrou recibo {mes}/{ano}'

    # tenta fazer o download do recibo em PDF
    res = session.get(url_base + str(download[1]), verify=False)
    if res.headers.get('content-type', '') != 'application/pdf':
        print('❌ Erro - Não pode fazer download do recibo')
        return 'Erro - Não pode fazer download do recibo'

    # salva o PDF do recibo
    nome = f'recibo_sn {cnpj} {mes}{ano}.pdf'
    _download_file(nome, res)

    print('✔ Recibo disponível')
    return 'Recibo disponível'


@_time_execution
def run():
    comp = _get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp:
        return False
    comp = comp.split('/')

    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    '''# inicia o driver para realizar o login no site
    driver = exec_phantomjs()
    if isinstance(driver, str):
        raise Exception(driver)'''

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    text = 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.'
    while text == 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.':
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--window-size=1920,1080')
    
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            cnpj, cpf, cod = empresa
    
            # printa o indice da empresa que está sendo executada
            _indice(count, total_empresas, empresa)
            
            captcha = 'erro'
            while captcha == 'erro':
                # loga no site do simples nacional com web driver e retorna uma sessão de request
                status, driver = _initialize_chrome(options)
                session = _new_session_sn(cnpj, cpf, cod, 'recibos', driver)
                
                if session == 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.':
                    print('❌ Erro Login - Caracteres anti-robô inválidos. Tente novamente')
                    captcha = 'erro'
                    driver.quit()
                else:
                    captcha = 'ok'
                    # se existe uma sessão realiza a consulta
                    if isinstance(session, Session):
                        text = consulta_recibos(cnpj, *comp, session)
                        driver.quit()
                    else:
                        text = session
                        driver.quit()
                    
                    # escreve na planilha de andamentos o resultado da execução atual
                    _escreve_relatorio_csv(f'{cnpj};{text}')
    return True


if __name__ == '__main__':
    run()
