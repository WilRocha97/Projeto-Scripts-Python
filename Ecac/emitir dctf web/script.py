# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from base64 import b64encode 
from requests import Session
from pyautogui import press
from ssl import CERT_NONE
from sys import path
import subprocess, websocket, random, json, re, os

path.append(r'..\..\_comum')
from pyautogui_comum import wait_img, get_comp
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, download_file, \
open_lista_dados, where_to_start, which_cert, content_or_soup, _headers


# variaveis globais
_json = 'data.json'
_url_base = 'https://dctfweb.cav.receita.fazenda.gov.br/AplicacoesWeb/DCTFWeb'

_url_default = f'{_url_base}/Default.aspx'
_url_lista = f'{_url_base}/Paginas/ListaDctfs.aspx'
_url_vinculacoes = f'{_url_base}/Paginas/ResumoVinculacoes/VinculacoesCompleta.aspx'


def get_data_json(arq, tag):
    with open(arq, encoding='utf8') as f:
        data = json.load(f)
        return data[tag]


def get_extra_data(content, aux, data):
    field = [ 
        'cphConteudo', f'TabelaVinculacoes{aux}',
        'GridViewVinculacoes', 'chkEmitirGuiaPagamento'
    ]

    soup = BeautifulSoup(content, 'html.parser')
    inputs = soup.findAll('input', id=re.compile('_'.join(field)))

    for inp in inputs:
        name = inp.get('name', '')
        if name: data[name] = 'on' 

    aux = 'ctl00$' + '$'.join(field[:-2]) + '$modalAbaterPagamentosAnteriores$txtNumeroGuiaAbater'
    if soup.find('input', attrs={'name':aux}): data[aux] = ''

    return data


@content_or_soup
def get_post_ecac(soup):
    infos = [
        soup.find('input', attrs={'id':'__VIEWSTATE'}),
        soup.find('input', attrs={'id':'__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id':'__EVENTVALIDATION'}),
    ]

    # state, generator, validation
    return tuple(info.get('value', '') for info in infos if info)


def get_tipo_doc(res):
    base = 'cphConteudo_TabelaVinculacoes@_pnVinculacoes'
    soup = BeautifulSoup(res.content, 'html.parser')

    for test in ('DARF', 'DAE'):
        aux = base.replace('@', test)
        if soup.find('div', attrs={'id': aux}): return test

    if 'Nenhuma declaração encontrada' in res.text: return ''

    print(soup.prettify())
    raise Exception('tipo documento não era DARF/DAE')


def ws_handler(auth_key, cookies):
    ws_key = str(b64encode(bytes([random.randint(0, 255) for _ in range(16)])), 'ascii')
    ws_msg = '{"acao" : "INICIAR_SESSAO", "sistemaOrigem" : "DctfWeb", "hmac" : @}'

    ws_headers = {
        'Sec-WebSocket-Key': ws_key, 'Sec-WebSocket-Version': '13',
        'Upgrade': 'websocket', 'User-Agent': _headers['User-Agent'],
    }

    url_ws = 'wss://assinadoc.estaleiro.serpro.gov.br/server/socket'
    cookies = "; ".join(f"{i}={j}" for i, j in cookies.items())

    ws = websocket.WebSocket(sslopt={"cert_reqs": CERT_NONE})
    ws.connect(url_ws, headers=ws_headers, cookie=cookies)
    ws.send(ws_msg.replace('@', auth_key))

    return ws


def transmite_dctf_web(cnpj, target, session):
    print('>>> transmitindo documento')

    url_assinadoc = f'{_url_base}/Paginas/Assinadoc/Assinadoc.aspx'
    url_download = 'https://assinadoc.estaleiro.serpro.gov.br/server/jnlp'

    res = session.get(_url_default, headers=_headers, verify=False)
    state, generator, validation = get_post_ecac(content=res.content)

    data = get_data_json(_json, 'lista')
    data['__VIEWSTATE'], data['__VIEWSTATEGENERATOR'] = state, generator
    data['__EVENTTARGET'], data['__EVENTVALIDATION'] = target, validation

    res = session.post(_url_lista, data, headers=_headers, verify=False)
    state, generator, validation = get_post_ecac(content=res.content)
    aux = get_tipo_doc(res)
    
    if not aux:
        return 'Nenhuma declaracao encontrada'

    data = get_data_json(_json, 'vinculacoes_transmitir')
    data = { k.replace('@', aux):v for k, v in data.items()}
    data['__VIEWSTATE'], data['__VIEWSTATEGENERATOR'] = state, generator
    data['__EVENTVALIDATION'] = validation

    session.post(_url_vinculacoes, data, headers=_headers, verify=False)
    res = session.get(url_assinadoc, headers=_headers, verify=False)
    state, generator, validation = get_post_ecac(content=res.content)

    d_info = {
        'doc': re.search(r"var documento = '(.*)';", res.text),
        'thumb': re.search(r"var thumbprint = '(.*)';", res.text),
        'auth_key': re.search(r"\"ChaveAutenticacao\":\"(.{64})\"};", res.text),
    }

    if not all(tuple(d_info.values())):
        return 'Erro não encontrou chave de autentificacao'

    d_info = {k:v.group(1).strip() for k, v in d_info.items()}
    try:
        ws = ws_handler(d_info['auth_key'], session.cookies.get_dict())
        ws_info = json.loads(ws.recv())
    except:
        return 'Erro assinador'

    data = get_data_json(_json, 'download')
    data['thumbprint'], data['documento'] = d_info['thumb'], d_info['doc']
    data['chaveAutenticacao'], data['sessionid'] = ws_info['chaveAutenticacao'], ws_info['sessionid']

    res = session.post(url_download, data, headers=_headers, verify=False)
    if res.headers.get('content-type', '') != 'application/x-java-jnlp-file':
        return 'Erro request download assinador'

    name = f'assinador {cnpj}.jnlp'
    if not download_file(name, res, r'ignore/temp'):
        return 'Erro download assinador'

    print('>>> assinando documento')
    with subprocess.Popen(['javaws', '-Xnosplash', f'ignore\\temp\\{name}']) as p:
        if not wait_img('btn_exec.png', conf=0.9):
            return 'Erro assinador não abriu'

        press('enter')
        if not wait_img('btn_success.png', conf=0.9):
            return 'Erro documento não assinado'

        press('enter')

    if p.poll(): p.terminate()
    ws_info = json.loads(ws.recv())
    ws.close()

    data = {
        '__EVENTTARGET': '', '__EVENTARGUMENT': '', '__VIEWSTATE': state,
        '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION': validation,
        'ctl00$cphConteudo$hdDocumentoAssinado': ws_info['documentoAssinado'],
        'ctl00$cphConteudo$hdHashDocumento': ws_info['chaveAutenticacao'],
        'ctl00$cphConteudo$hdChaveAutenticacao': ws_info['chaveAutenticacao'],
    }

    res = session.post(url_assinadoc, data, headers=_headers, verify=False)
    os.remove(os.path.join('ignore', 'temp', name))
    return True


def emite_dctf_web(cnpj, session, target=None, comp=None):
    print('>>> emitindo documento')
    url_download = f'{_url_base}/Paginas/GuiaPagamento/DownloadPdfGuiaPagamento.aspx'

    res = session.get(_url_default, headers=_headers, verify=False)
    state, generator, validation = get_post_ecac(content=res.content)
    if not target:
        target = get_ev_target(res.content, comp, session)

    data = get_data_json(_json, 'lista')
    data['__VIEWSTATE'], data['__VIEWSTATEGENERATOR'] = state, generator
    data['__EVENTTARGET'], data['__EVENTVALIDATION'] = target, validation

    res = session.post(_url_lista, data, headers=_headers, verify=False)
    state, generator, validation = get_post_ecac(content=res.content)
    aux = get_tipo_doc(res)
    
    if not aux:
        return 'Nenhuma declaracao encontrada'

    data = get_data_json(_json, 'vinculacoes_emitir')
    data = get_extra_data(res.content, aux, data)
    data = { k.replace('@', aux):v.replace('@', aux) for k, v in data.items()}
    data['__VIEWSTATE'], data['__VIEWSTATEGENERATOR'] = state, generator
    data['__EVENTVALIDATION'] = validation

    valor = BeautifulSoup(res.content, 'html.parser').find('span', attrs={
        'id': f'cphConteudo_TabelaVinculacoes{aux}_GridViewVinculacoes_lblValorSaldoPagar_0'
    }).text

    session.post(_url_vinculacoes, data, headers=_headers, verify=False)
    res = session.get(url_download, headers=_headers, verify=False)

    if res.headers.get('content-type', '') != 'application/octet-stream':
        return f'transmitido - Erro request download guia;{valor}'

    name = f'dctf-web {cnpj}.pdf'        
    if not download_file(name, res):
        return f'transmitido - Erro download guia;{valor}'

    return f'emitido;{valor}'


def get_ev_target(content, comp, session):
    state, generator, validation = get_post_ecac(content=content)

    data = get_data_json(_json, 'lista')
    data['__VIEWSTATE'], data['__VIEWSTATEGENERATOR'] = state, generator
    data['__EVENTVALIDATION'] = validation

    res = session.post(_url_lista, data, headers=_headers, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')

    table_id = 'ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs'
    table = soup.find('table', attrs={'id': table_id})
    table = table.findAll('tr', recursive=False)

    if len(table) == 1: return ''
    not_allowed = ('original sem movimento', 'original zerada', 'retificadora')

    for tr in table[1:]:
        tds = tr.findAll(recursive=False)
        if not tds: continue

        if tds[0].text.strip().lower() != comp:
            continue

        tipo = tds[4].text.strip().lower()

        if tipo == 'original':
            return tds[-1].a.get('id').replace('_', '$')

        if tipo in not_allowed:
            return tipo

        raise Exception('Erro target - avaliar manualmente')

    return ''


def analisa_dctf_web(cnpj, comp, session):    
    res = session.get(_url_default, headers=_headers, verify=False)
    op = get_ev_target(res.content, comp, session)

    if 'Editar' in op:
        result = transmite_dctf_web(cnpj, op, session)

        if isinstance(result, str):
            return result

        return emite_dctf_web(cnpj, session, comp=comp)

    elif 'Visualizar' in op:
        return emite_dctf_web(cnpj, session, target=op)

    elif '' == op:
        return f'Nenhuma declaração de {comp} encontrada'

    return op


@time_execution
def run():
    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas:
        return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp:
        return False

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
            try:
                text = analisa_dctf_web(cnpj, comp, session)
            except ConnectionError:
                text = f'{cnpj};Erro Login - {res["value"]}'
                print('\t--connection error')
        else:
            text = f'Erro Login - {res["value"]}'
            print(text)

        print(cnpj, text)
        escreve_relatorio_csv(f"{cnpj};{text}")

    if isinstance(session, Session):
        session.close()
    return True


if __name__ == '__main__':
    run()
