# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from sys import path
import os

path.append(r'..\..\_comum')
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, open_lista_dados, where_to_start, which_cert, _headers, _indice


# variaveis globais
exec_path = r'execucao/documentos'


def download_message(name, content, css):
    os.makedirs(exec_path, exist_ok=True)

    soup = BeautifulSoup(content, 'html.parser')
    tag = soup.new_tag('style', type="text/css")
    tag.string = css
    
    soup = BeautifulSoup(content, 'html.parser')
    soup.find('link').append(tag)

    name = os.path.join(exec_path, name)
    with open(name, 'w', encoding="latin-1", errors='ignore') as arq:
        arq.write(soup.prettify())

    return True


def procurar_mensagen(cnpj, s):
    print('>>> Procurando mensagens')

    url_base = 'https://cav.receita.fazenda.gov.br'
    url = f'{url_base}/Servicos/ATSDR/CaixaPostal.app/Action'
    url_css = f'{url_base}/Servicos/ATSDR/CaixaPostal.app/Content/cxpostal.css'

    res = s.get(f'{url}/ListarMensagemAction.aspx', headers=_headers, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')

    msgs = soup.find('table', attrs={'id': 'gridMensagens'})
    if not msgs:
        return 'Não tem mensagens'
    msgs = len(msgs.findAll('tr'))
    
    css = s.get(url_css, headers=_headers)

    mensagem = False
    for linha in range(2, msgs):
        id_parc = 'gridMensagens_ctl' + str(linha).rjust(2, '0')

        objeto = soup.find('img', attrs={'id': f'{id_parc}_imgRelevancia'})
        if not objeto:
            continue

        lido = soup.find('a', attrs={'id': f'{id_parc}_inLeitura'})
        if not lido:
            continue

        conds = (
            'imagens/exclamation.gif' == objeto.get('src', ''),
            'imagens/aaMsgNaoLida.gif' == lido.find('img').get('src', '')
        )
        if not all(conds):
            continue

        objeto_m = soup.find('a', attrs={'id': f'{id_parc}_nmAssunto'})
        if not objeto_m:
            continue
        objeto_m = objeto_m.get('href')

        res = s.get(f"{url}/{objeto_m}", headers=_headers)
        name = f'{cnpj} - mensagem {linha-1}.html'
        
        if not download_message(name, res.content, css.text):
            continue

        mensagem = True

    if mensagem:
        print('❗ Mensagens salvas')
        return 'Mensagens salvas'
    else:
        print('✔ Não tem mensagens')
        return 'Não tem mensagens'


@time_execution
def run():
    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas:
        return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cert = empresa

        _indice(count, total_empresas, empresa)

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
            text = procurar_mensagen(cnpj, session)
        else:
            text = f'Erro Login - {res["value"]}'
            print(text)
        escreve_relatorio_csv(f'{cnpj};{text}')

    if isinstance(session, Session):
        session.close()
    return True


if __name__ == '__main__':
    run()
