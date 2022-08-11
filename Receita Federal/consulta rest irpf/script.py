# -*- coding: utf-8 -*-
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from pathlib import Path
from sys import path
import json, os, re

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from captcha_comum import break_normal_captcha
from selenium_comum import exec_phantomjs, get_elem_img
from comum_comum import time_execution, escreve_relatorio_csv, \
open_lista_dados, where_to_start


# variaveis globais
_url_base = 'http://servicos.receita.fazenda.gov.br/servicos/consrest/atual.app/paginas'


def new_session_rest(cpf, ano, driver):
    print('\n>>> nova sessao', cpf, 'ano', ano)    
    driver.get(f'{_url_base}/mobile/restituicaomobi.asp')

    for i in range(20):
        try:
            img = driver.find_element_by_id('imgCaptcha')
            break
        except:
            sleep(1)
    else:
        return 'Erro captcha - Não encontrou captcha'

    img = get_elem_img(img, driver)
    captcha = break_normal_captcha(img)
    if isinstance(captcha, str): return captcha

    session = Session()
    for c in driver.get_cookies():
        session.cookies.set(c['name'], c['value'])

    driver.delete_all_cookies()
    return captcha, session


def save_html(cpf, ano, data):
    data = {k:BeautifulSoup(y, 'html.parser') for k, y in data.items()}
    now = datetime.now().strftime('%d/%m/%Y - %H:%M:%S')

    with open('base.html', 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    soup.find('span', attrs={'id':'txt_ano'}).append(ano)
    soup.find('span', attrs={'id':'txt_now'}).append(now)
    soup.find('span', attrs={'id':'txt_msgStatus'}).append(data['msgStatus'])
    soup.find('span', attrs={'id':'txt_cpfFormatado'}).append(data['cpfFormatado'])
    soup.find('span', attrs={'id':'txt_nomeContribuinte'}).append(data['nomeContribuinte'])

    labels = ('resultado', 'msgInfo', 'observacao', 'observacao2', 'tipoDecl')
    for label in labels:
        if not data[label]: continue
        soup.find('span', attrs={'id': f'txt_{label}'}).append(data[label])

    if data['codSituacao'].text.strip() == '2':
        soup.find('span', attrs={'id':'txt_lote'}).append(data['lote'])
        soup.find('span', attrs={'id':'txt_banco'}).append(data['banco'])
        soup.find('span', attrs={'id':'txt_disponivelEm'}).append(data['disponivelEm'])
        if data['agencia']:
            soup.find('span', attrs={'id':'txt_agencia'}).append(data['agencia'])
        
        div = soup.find('div', attrs={'id':'div_banco'})
        div['style'] = 'display:block;'

    e_dir = Path('execucao', 'docs')
    os.makedirs(e_dir, exist_ok=True)

    with open((e_dir / f'{cpf}-{ano}.html'), 'w') as f:
        f.write(soup.prettify())


def normalize_text(data):
    result = [data['nomeContribuinte']]
    if 'liberação de sua restituição' in data['msgStatus']:
        result.append('Restituição liberada')
    else:
        result.append(data['msgStatus'].strip().replace('.', ''))

    if data['msgInfo']:
        aux = data['msgInfo'].split(':')[-1].replace('</b>', '')
        result.append(aux.replace('&ordf;', 'a').strip())

    if data['codSituacao'].strip() == '2':
        result.append(data['banco'].strip())
        if data['agencia']:
            result.append(data['agencia'].strip())
        result.append(data['lote'].strip())
        result.append(data['disponivelEm'].strip())

    return ';'.join(result)


def consulta_rest(cpf, nasc, ano, captcha, session):
    print('>>> consultando restituição')

    url_login = f'{_url_base}/util/tratamentoCaptchaB.asp'
    url_query = f'{_url_base}/webservice/restituicaoJson.asp'

    data = {
        'cpf': cpf, 'data_nascimento': nasc, 'exercicio': ano,
        'txtTexto_captcha_serpro_gov_br': captcha['code']
    }

    res = session.post(url_login, data, verify=False)
    if 'mensagemCaptcha:sucesso' not in res.text:
        return 'Erro captcha - ' + res.text.split(':', 1)[-1]

    code = ''.join(i for i in res.text if i.isdigit())
    query = f'cpf={cpf}&exercicio={ano}&hash={code}&data_nascimento={nasc}'

    # precisa do soup para resolver problemas de encoding no response
    res = session.post(f'{url_query}?{query}', verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')

    regex = re.compile(r'([^{,:])"([^,:}])', flags=re.IGNORECASE)
    regex_lambda = lambda x: f"{x.group(1)}'{x.group(2)}"

    data = regex.sub(regex_lambda, soup.text)
    data = json.loads(data)

    try:
        save_html(cpf, ano, data)
    except Exception as e:
        raise Exception(str(e), data)

    session.close()
    return normalize_text(data)


@time_execution
def run():
    empresas = open_lista_dados()
    if not empresas: return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None: return False

    ano = get_comp(printable='yyyy', strptime='%Y')
    if not ano: return False

    driver = exec_phantomjs()
    if isinstance(driver, str):
        raise Exception(driver)

    for cpf, dt_nasc in empresas[index:]:
        res = new_session_rest(cpf, ano, driver)
        if not isinstance(res, str):
            text = consulta_rest(cpf, dt_nasc, ano, *res)
        else: 
            text = res

        text = f'{cpf};{ano};{text}'

        print('>>>', text)
        escreve_relatorio_csv(text, nome=f'resumo{ano}')
    
    return True


if __name__ == '__main__':
    run()