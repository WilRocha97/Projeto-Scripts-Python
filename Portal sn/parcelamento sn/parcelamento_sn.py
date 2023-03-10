# -*- coding: utf-8 -*-
import time

from pyautogui import confirm
from selenium import webdriver
from bs4 import BeautifulSoup
from sys import path
import sys

path.append(r'..\..\_comum')
from sn_comum import new_session_sn
from pyautogui_comum import get_comp
from chrome_comum import initialize_chrome
from comum_comum import time_execution, download_file, escreve_relatorio_csv, open_lista_dados, where_to_start, content_or_soup, _indice


@content_or_soup
def get_info_post(soup):
    infos = [
        soup.find('input', attrs={'id': '__VIEWSTATE'}),
        soup.find('input', attrs={'id': '__VIEWSTATEGENERATOR'}),
    ]

    # state, generator
    return tuple(info.get('value', '') for info in infos if info)


def consulta_parcelamento(cnpj, tipo, comp, session):
    url_base = 'https://www8.receita.fazenda.gov.br'
    url_home = f'{url_base}/SimplesNacional/Aplicacoes/ATSPO/{tipo}.app'
    
    try:
        print('>>> Consultando parcelamento')
        res = session.get(url_home, verify=False)
        
        
        state, generator = get_info_post(content=res.content)
    except:
        sys.stdout.write('❌ ')
        text = 'Erro ao consultar parcelamento'
        return text
            
    data = {
        '__EVENTTARGET': 'ctl00$contentPlaceH$linkButtonEmitirDAS',
        '__EVENTARGUMENT': '',
        '__VIEWSTATE': state,
        '__VIEWSTATEGENERATOR': generator
    }

    res = session.post(f'{url_home}/default.aspx', data, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')
    
    try:
        state, generator = get_info_post(soup=soup)
    except:
        return 'Erro no site'

    if 'Não há parcelamento' in res.text:
        sys.stdout.write('>>> ')
        return f'Não há parcelamento {tipo}'
    
    if 'Não há parcela disponível para reimpressão' in res.text:
        sys.stdout.write(' ')
        return f'Não há parcela {tipo} disponível para reimpressão'
    
    if 'Não há parcela disponível para impressão' in res.text:
        sys.stdout.write(' ')
        return f'Não há parcela {tipo} disponível para impressão'
    
    if 'Há um pedido de parcelamento para o contribuinte aguardando confirmação do pagamento da primeira parcela.' in res.text:
        sys.stdout.write('❌ ')
        return 'Há um pedido de parcelamento para o contribuinte aguardando confirmação do pagamento da primeira parcela'
    
    if 'Há um pedido de parcelamento para o contribuinte com primeira parcela ainda não vencida.' in res.text:
        sys.stdout.write('❌ ')
        return 'Há um pedido de parcelamento para o contribuinte com primeira parcela ainda não vencida.'
    
    for i in ('', 's'):
        id_table = 'ctl00_contentPlaceH_gdvParcela' + i
        if not soup.find('table', attrs={'id': id_table}):
            continue

        break

    data = {
        '__EVENTTARGET': '', '__EVENTARGUMENT': '',
        '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
        'ctl00$contentPlaceH$btnContinuar': 'Continuar'
    }

    action = soup.find('form', attrs={'name': 'aspnetForm'}).get('action')
    res = session.post(f'{url_home}/{action}', data, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')
    state, generator = get_info_post(soup=soup)

    table = soup.find('table', attrs={'id': id_table})
    parcs = tuple((i.find('td'), i.find('a')) for i in table.findAll('tr')[1:])

    atraso = []
    for date, target in parcs:
        date = date.text.strip()
        
        if comp != date:
            atraso.append(date)
            continue

        data = {
            '__EVENTTARGET': target.get('id', '').replace('_', '$'),
            '__EVENTARGUMENT': '', '__VIEWSTATE': state,
            '__VIEWSTATEGENERATOR': generator
        }

    res = session.post(f'{url_home}/{action}', data, verify=False)
    if res.headers.get('content-type', '') != 'application/octet-stream':
        sys.stdout.write('❌ ')
        return f'Erro - {tipo} indisponível;{" ".join(atraso)}'

    comp = comp.replace("/", "")
    download_file(f'parc-{tipo} {cnpj} {comp}.pdf', res)

    sys.stdout.write('✔ ')
    return f'parc. {tipo} {comp} disponível;{" ".join(atraso)}'


def get_tipo():
    tipos = {'pert/pertsn': 'pertsn', 'das/snparc': 'snparc', 'mei/parcmei': 'parcmei'}

    text = 'Quais parcelamentos serão gerados?'
    tipo = confirm(title='Tipo de parcelamento', text=text, buttons=tuple(tipos.keys()))
    if tipo is None:
        return False

    return tipos[tipo]


@time_execution
def run():
    tipo = get_tipo()
    if not tipo:
        return False

    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp:
        return False

    empresas = open_lista_dados()
    if not empresas:
        return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    status, driver = initialize_chrome(options)
    if isinstance(driver, str):
        raise Exception(driver)
    
    resumo = f'resumo {comp.replace("/", "")} {tipo}'
    print('>>> Executando para', tipo, comp, f"salvando em '{resumo}.csv'\n")

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cpf, cod = empresa

        _indice(count, total_empresas, empresa)

        text = 'Erro ao consultar parcelamento'
        session = 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.'
        while session == 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.' or text == 'Erro ao consultar parcelamento' or text == f'Não carregou tabela {tipo}':
            session = new_session_sn(cnpj, cpf, cod, tipo, driver)
            
            if isinstance(session, str):
                if not session == 'Erro Login - Caracteres anti-robô inválidos. Tente novamente.':
                    escreve_relatorio_csv(f'{cnpj};{session}', nome=resumo)
                print('>>>', session)
                text = None
            else:
                try:
                    text = consulta_parcelamento(cnpj, tipo, comp, session)
                except:
                    text = 'Erro ao consultar parcelamento'
                    driver.quit()
                    time.sleep(1)
                    
                    status, driver = initialize_chrome(options)
                    if isinstance(driver, str):
                        raise Exception(driver)
                    
                if text == 'Erro ao consultar parcelamento':
                    print(text)
                elif text ==  'Não carregou tabela snparc':
                    print(text)
                else:
                    print(text)
                    escreve_relatorio_csv(f'{cnpj};{text}', nome=resumo)

                session.close()

    driver.quit()
        
    return True


if __name__ == '__main__':
    run()
