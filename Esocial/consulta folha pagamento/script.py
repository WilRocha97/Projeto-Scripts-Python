# -*- coding: utf-8 -*-
from pyautogui import confirm
from bs4 import BeautifulSoup
from requests import Session
from sys import path
import json

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from esocial_comum import new_session_esocial, new_login_esocial
from comum_comum import time_execution, escreve_relatorio_csv, escreve_header_csv, \
open_lista_dados, where_to_start, which_cert, _headers


def verifica_meses(meses, anos, session):
    url_base = 'https://www.esocial.gov.br'
    url_home = f'{url_base}/portal/Home/Inicial?tipoEmpregador=EMPREGADOR_GERAL'
    url_mesf = f'{url_base}/portal/FolhaPagamento/GestaoFolha/ObterMeses?ano=@ano'
    url_sitf = f'{url_base}/portal/FolhaPagamento/GestaoFolha/ObterSituacaoFolha?ano=@ano&mes=@mes'

    res = session.get(url_home, headers=_headers, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')
    compl = soup.find('a', attrs={'id': 'menuGestaoFolha'}).get('href', '')

    res = session.get(f'{url_base}{compl}', headers=_headers, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')
    
    if not anos:
        anos = tuple(map(lambda x: x.text.strip(), soup.findAll('a', attrs={'class':'seletorAno'})))

    data = {}
    for ano in anos:
        url_aux = url_mesf.replace('@ano', ano)
        res = session.get(url_aux, headers=_headers, verify=False)
        data[ano] = json.loads(res.text)

    aberta, fechada = [], []
    for ano, value in data.items():
        if not meses:
            meses = tuple(map(lambda x: x['Valor'], value['Objeto']['Itens']))

        for mes in meses:
            url_aux = url_sitf.replace('@ano', ano).replace('@mes', mes)
            res = session.get(url_aux, headers=_headers, verify=False)
            info = json.loads(res.text)

            if info['Objeto']['NomeSituacaoFolha'].lower() == 'aberta':
                aberta.append(f'{mes}/{ano}')
            else:
                fechada.append(f'{mes}/{ano}')

    return ' '.join(aberta) + ';' + ' '.join(fechada)


@time_execution
def run():
    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas: return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None: return False

    res = confirm('Pesquisar mês especifico?', buttons=('sim', 'não'))
    if not res: return False

    if res == 'sim':
        comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
        if not comp: return False

        print('>>> consultando', comp)
        comp = tuple(tuple(i) for i in comp.split('/'))
    else:
        print('>>> consultando todos os meses')
        comp = (None, None)

    for cnpj, cert in empresas[index:]:
        if prev_cert != cert:
            prev_cert = cert
            try:
                session = session.close()
            except AttributeError:
                pass

        if not isinstance(session, Session):
            aux = which_cert(cert)
            if isinstance(aux, str):
                raise Exception(aux)

            session = new_session_esocial(*aux)
            if isinstance(session, str):
                raise Exception(session)

        res = new_login_esocial(cnpj, session)
        if res == '':
            text = verifica_meses(*comp, session)
            print('consultado com sucesso')
        else:
            text = res
            print(text)

        escreve_relatorio_csv(f"{cnpj};{text}")

    escreve_header_csv('cnpj;meses_abertos;meses_fechados')
    if isinstance(session, Session): session.close()
    
    return True


if __name__ == '__main__':
    run()