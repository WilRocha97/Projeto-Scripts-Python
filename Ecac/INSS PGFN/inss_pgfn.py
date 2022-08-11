# -*- coding: utf-8 -*-
from requests.exceptions import ConnectionError
from datetime import date
from bs4 import BeautifulSoup
from requests import Session
from pyautogui import prompt
from sys import path
import sys, time

path.append(r'..\..\_comum')
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, _escreve_header_csv, open_lista_dados, where_to_start, which_cert, _headers, _indice
from comum import atualiza_info, salvar_arquivo


# variaveis globais
exec_path = r'execucao/docs'


def atualiza_token_view(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')

    try:
        token = soup.find(attrs={"name": "DTPINFRA_TOKEN"})['value']
        # token = find_token['value']
    except:
        token = soup.find('input', attrs={'name': 'DTPINFRA_TOKEN'}).get('value')

    view = soup.find('input', attrs={'id': 'javax.faces.ViewState'}).get('value')
    return token, view


def verifica_avisos(pagina, div_class='hd'):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    try:
        div = soup.find('div', attrs={'class': div_class})
        msg = div.find('li', attrs={'class': 'message'}).text
        return msg
    except:
        return False


def lista_acordos(pagina):
    situacoes_parc = [
        'ATIVO (EM DIA)',
        'ATIVO (EM ATRASO)',
    ]

    soup = BeautifulSoup(pagina.content, 'html.parser')
    tabela = soup.find('tbody', attrs={'id': 'formComposicaoPedido:j_id36:tbody_element'})
    if not tabela:
        return []

    lista = []
    for linha in tabela.findAll('tr'):
        colunas = linha.findAll('td')

        if colunas[3].text.strip() not in situacoes_parc:
            continue
        
        acordo = colunas[0].text
        link = colunas[-1].find('input').get('name')
        lista.append((acordo, link))

    return lista


def download_guia(pagina, venc):
    info, atrasados = False, []
    day, month, year = [int(i) for i in venc.split('/')]
    venc = date(year, month, day)

    soup = BeautifulSoup(pagina.content, 'html.parser')
    tabela = soup.find('tbody', attrs={'id': 'formDetalharExtratoParcelamento:j_id92:tbody_element'})
    for linha in tabela.findAll('tr'):
        colunas = linha.findAll('td')

        day, month, year = [int(i) for i in colunas[1].text.strip().split('/')]
        td_venc = date(year, month, day)

        if td_venc < venc:
            dt_pagamento = colunas[3].text
            if dt_pagamento == '-':
                atrasados.append(td_venc.strftime('%d/%m/%Y'))
            continue

        if td_venc != venc:
            break
        token, view = atualiza_token_view(pagina)
        try:
            target = colunas[8].find('input').get('name')
        except:
            break
            
        info = {
            'formDetalharExtratoParcelamento': 'formDetalharExtratoParcelamento',
            'DTPINFRA_TOKEN': token, target: '', 'javax.faces.ViewState': view
        }
        break

    return info, atrasados


def verifica_parcelamento(cnpj, session, venc):
    '''headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }'''

    url_home = 'https://cav.receita.fazenda.gov.br/ecac/'
    url_base = 'https://www2sp.dataprev.gov.br/ParcWebPrevPGFNInternet'

    url_inter = 'https://cav.receita.fazenda.gov.br/Servicos/ATBHE/DATAPREV/PARCWEBPREVPGFN/App.aspx'
    url_login = 'https://www2.dataprev.gov.br/Ecac/LoginServlet'

    url_dtpr = f'{url_base}/LoginServlet'  # post da url_dtpr visivel somente pelo chrome
    url_parc = f'{url_base}/pages/index.xhtml'
    url_acordo = f'{url_base}/pages/pedido/extratoParcelamento.xhtml'
    url_download = f'{url_base}/pages/pedido/detalharExtratoParcelamento.xhtml'

    session.get(url_home, headers=_headers, verify=False)
    pagina = session.get(url_inter, headers=_headers, verify=False)

    print('>>> Acessando dataprev')
    state, generator, validation, msg64 = atualiza_info(pagina)
    info = {
        '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
        '__EVENTVALIDATION': validation, 'mensagemComBase64': msg64
    }
    
    # Necessário um loop pois o site esta bloqueando o 1º acesso a página
    for i in range(4):
        pagina = session.post(url_login, info, headers=_headers, verify=False)
        soup = BeautifulSoup(pagina.content, 'html.parser')

        msg = soup.find('input', attrs={'name': 'mensagem'}).get('value')
        timestamp = soup.find('input', attrs={'name': 'timestamp'}).get('value')
        signature = soup.find('input', attrs={'name': 'signature'}).get('value')
        
        info = {
            'mensagemComBase64': msg64, 'mensagem': msg, 'service': 'parcweb',
            'version': '1.0', 'accessKey': 'parcweb', 'associationId': 'parcweb',
            'protocol': '2', 'timestamp': timestamp, 'signature': signature
        }

        pagina = session.post(url_dtpr, info, headers=_headers, verify=False)
        time.sleep(2)
        if not verifica_avisos(pagina, 'bd'):
            break
    else:
        sys.stdout.write('>>> ')
        return 'Acesso negado para a lista de parcelamentos'

    token, view = atualiza_token_view(pagina)
    info = {
        'form': 'form', 'DTPINFRA_TOKEN': token, 'javax.faces.ViewState': view,
        'form:extratoParcelamento': 'form:extratoParcelamento'
    }

    pagina = session.post(url_parc, info, headers=_headers, verify=False)
    aviso = verifica_avisos(pagina)
    if aviso:
        sys.stdout.write('>>> ')
        return aviso

    token, view = atualiza_token_view(pagina)
    acordos = lista_acordos(pagina)
    if not acordos:
        sys.stdout.write('>>> ')
        return 'Empresa sem acordos ativos.'

    textos = []
    for acordo, link in acordos:
        print('>>> Verificando acordo', acordo)
        info = {
            'DTPINFRA_TOKEN': token,
            'formComposicaoPedido': 'formComposicaoPedido',
            link: '', 'javax.faces.ViewState': view
        }

        pagina = session.post(url_acordo, info, headers=_headers, verify=False)
        guia, atrasados = download_guia(pagina, venc)

        sit_guia = 'Guia atual não disponivel'
        if guia:
            salvar = session.post(url_download, guia, headers=_headers, verify=False)
            str_venc = venc.replace('/', '-')
            list_venc = venc.split('/')

            nome = ( 
                f'{cnpj};IMP_FISCAL;{str_venc};Parcelamento Previdenciário PGFN' +
                f'{list_venc[0]};{list_venc[1]};Parcelamento;INSS 2 - {acordo}.pdf'
            )
            salvar_arquivo(nome, salvar)
            sit_guia = 'Guia atual disponivel'
            sys.stdout.write('✔ ')

        sit_atrasados = 'Sem guias em atraso'
        if atrasados:
            sit_atrasados = ', '.join(atrasados)
            sys.stdout.write('>>> ')

        textos.append(f'{acordo};{sit_guia};{sit_atrasados}')

    return '\n'.join(textos)


@time_execution
def run():
    print("Insira a data de vencimento na caixa de texto... ", end='')

    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas:
        return False

    venc = prompt("Digite o vencimento das guias 00/00/0000:")
    if not venc:
        print("Execução cancelada")
        return False

    print(venc)
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
            try:
                text = verifica_parcelamento(cnpj, session, venc)
                print(text)
            except ConnectionError:
                text = f'Erro Login - {res["value"]}'
                print('\t--connection error')
        else:
            text = f'Erro Login - {res["value"]}'
            print(text)

        escreve_relatorio_csv(f'{cnpj};{text}')

    if session is not None:
        session.close()

    header = "cnpj;acordo;sit. guia do mes atual;sit. guias em atraso"
    _escreve_header_csv(header)


if __name__ == '__main__':
    run()
