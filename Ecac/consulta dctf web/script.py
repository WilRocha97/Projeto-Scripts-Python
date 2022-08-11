# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, open_lista_dados, \
where_to_start, which_cert, content_or_soup, _headers


# variaveis globais
_url_base = 'https://dctfweb.cav.receita.fazenda.gov.br/AplicacoesWeb/DCTFWeb'
_url_default = f'{_url_base}/Default.aspx'


def consulta_dctf_web(cnpj, comp, session):
    print('>>> analisando tabela')

    res = session.get(_url_default, headers=_headers, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')

    table_id = 'ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs'
    table = soup.find('table', attrs={'id': table_id})
    table = table.findAll('tr', recursive=False)

    if len(table) == 1:
        return f'{cnpj};Nenhuma declaração de {comp} encontrada'

    text = []
    for tr in table[1:]:
        tds = tr.findAll(recursive=False)
        if not tds: continue

        if tds[0].text.strip().lower() != comp:
            continue

        text.append(f'{cnpj};' + ';'.join(i.text.strip().lower() for i in tds[:7]))

    return '\n'.join(text)


@time_execution
def run():
    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas: return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None: return False
    
    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp: return False

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

            session = new_session_ecac(*aux)
            if isinstance(session, str):
                raise Exception(session)

        res = new_login_ecac(cnpj, session)
        if res['key']:
            text = consulta_dctf_web(cnpj, comp, session)
            print('consultado com sucesso')
        else:
            text = f'{cnpj};Erro Login - {res["value"]}'
            print(text)

        escreve_relatorio_csv(text)

    if isinstance(session, Session): session.close()
    return True


if __name__ == '__main__':
    run()