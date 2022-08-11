# -*- coding: utf-8 -*-
from pyautogui import confirm
from bs4 import BeautifulSoup
from requests import Session
from datetime import date
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, escreve_header_csv, \
open_lista_dados, where_to_start, which_cert, content_or_soup, _headers


@content_or_soup
def atualiza_info(soup):

    infos = [
        soup.find('input', attrs={'id':'__VIEWSTATE'}),
        soup.find('input', attrs={'id':'__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id':'__EVENTVALIDATION'}),
        soup.find('input', attrs={'id':'mensagemComBase64'}),
    ]

    result = []
    for info in infos:
        try:
            result.append(info.get('value'))
        except:
            result.append('')

    return tuple(result) # state, generator, validation, msg64


def atualiza_token_view(content):
    soup = BeautifulSoup(content , 'html.parser')
    token = soup.find('input', attrs={'name':'DTPINFRA_TOKEN'}).get('value')
    view = soup.find('input', attrs={'id':'javax.faces.ViewState'}).get('value')
    return (token, view)


def verifica_avisos(content, div_class='hd'):
    soup = BeautifulSoup(content, 'html.parser')
    try:
        div = soup.find('div', attrs={'class': div_class})
        msg = div.find('li', attrs={'class': 'message'}).text
        return msg
    except:
        return False


def lista_acordos(content, ntab):
    situacoes_parc = ('ATIVO (EM DIA)', 'ATIVO (EM ATRASO)')

    soup = BeautifulSoup(content, 'html.parser')
    tabela = soup.find('tbody', id=f'formComposicaoPedido:j_id{ntab}:tbody_element')
    if not tabela: return []

    lista = []
    for linha in tabela.findAll('tr'):
        colunas = linha.findAll('td')

        if colunas[3].text.strip() not in situacoes_parc:
            continue
        
        acordo = colunas[0].text
        link = colunas[-1].find('input').get('name')
        lista.append((acordo, link))

    return lista


def download_guia(content, ntab, venc):
    info, atrasados = False, []
    day, month, year = [int(i) for i in venc.split('/')]
    venc = date(year, month, day)

    soup = BeautifulSoup(content, 'html.parser')
    tabela = soup.find('tbody', id=f'formDetalharExtratoParcelamento:j_id{ntab}:tbody_element')
    for linha in tabela.findAll('tr'):
        colunas = linha.findAll('td')

        day, month, year = [int(i) for i in colunas[1].text.strip().split('/')]
        td_venc = date(year, month, day)

        if td_venc < venc:
            if colunas[3].text != '-': continue
            atrasados.append(td_venc.strftime('%d/%m/%Y'))

        if td_venc != venc: break
        token, view = atualiza_token_view(content)
        try:
            target = colunas[8].find('input').get('name')
        except:
            break
            
        info = {
            'formDetalharExtratoParcelamento': 'formDetalharExtratoParcelamento',
            'DTPINFRA_TOKEN': token, target: '', 'javax.faces.ViewState': view
        }
        break

    return (info, atrasados)


def consulta_parcelamento(cnpj, tipo, venc, session):
    url_home = 'https://cav.receita.fazenda.gov.br/ecac/'
    url_base = f'https://www2sp.dataprev.gov.br/ParcWebPrev{tipo[0].upper()}Internet'

    url_inter = f'https://cav.receita.fazenda.gov.br/Servicos/ATBHE/DATAPREV/PARCWEBPREV{tipo[0].upper()}/App.aspx'
    url_login = 'https://www2.dataprev.gov.br/Ecac/LoginServlet'

    url_dtpr = f'{url_base}/LoginServlet'
    # url_dtpr visivel somente pelo chrome

    url_parc = f'{url_base}/pages/index.xhtml'
    url_acordo = f'{url_base}/pages/pedido/extratoParcelamento.xhtml'
    url_download = f'{url_base}/pages/pedido/detalharExtratoParcelamento.xhtml'

    session.get(url_home, headers=_headers, verify=False)
    res = session.get(url_inter, headers=_headers, verify=False)

    print('>>> Acessando dataprev')
    state, generator, validation, msg64 = atualiza_info(content=res.content)
    info = {
        '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
        '__EVENTVALIDATION': validation, 'mensagemComBase64': msg64
    }
    
    # Necessário um loop pois o site esta bloqueando o 1º acesso a página
    for i in range(4):
        res = session.post(url_login, info, headers=_headers, verify=False)
        soup = BeautifulSoup(res.content, 'html.parser')

        msg = soup.find('input', attrs={'name':'mensagem'}).get('value')
        time = soup.find('input', attrs={'name':'timestamp'}).get('value')
        signature = soup.find('input', attrs={'name':'signature'}).get('value')
        
        info = {
            'mensagemComBase64': msg64, 'mensagem': msg, 'service': 'parcweb',
            'version': '1.0', 'accessKey': 'parcweb', 'associationId': 'parcweb',
            'protocol': '2', 'timestamp': time, 'signature': signature
        }

        res = session.post(url_dtpr, info, headers=_headers, verify=False)
        if not verifica_avisos(res.content, 'bd'): break
    else:
        return f'{cnpj};Acesso negado para a lista de parcelamentos'

    token, view = atualiza_token_view(res.content)
    info = {
        'form': 'form', 'DTPINFRA_TOKEN': token, 'javax.faces.ViewState': view,
        'form:extratoParcelamento': 'form:extratoParcelamento'
    }

    res = session.post(url_parc, info, headers=_headers, verify=False)
    aviso = verifica_avisos(res.content)
    if aviso: return f'{cnpj};{aviso}'

    token, view = atualiza_token_view(res.content)
    acordos = lista_acordos(res, tipo[1])
    if not acordos: return f'{cnpj};Empresa sem acordos ativos.'

    textos = []
    for acordo, link in acordos:
        print('>>> verificando acordo', acordo)
        info = {
            'DTPINFRA_TOKEN': token, 'formComposicaoPedido': 'formComposicaoPedido',
            link: '', 'javax.faces.ViewState': view
        }

        res = session.post(url_acordo, info, headers=_headers, verify=False)
        guia, atrasados = download_guia(res, tipo[2], venc)

        sit_guia = 'Guia atual não disponivel'
        if guia:
            res = session.post(url_download, guia, headers=_headers, verify=False)
            venc = venc.replace('/', '')

            nome = f'parc. inss {tipo[0]}-{cnpj}-{acordo}-{venc}.pdf'
            download_file(nome, res)
            sit_guia = 'Guia atual disponivel'

        sit_atrasados = 'Sem guias em atraso'
        if atrasados:
            sit_atrasados = ', '.join(atrasados)

        textos.append(f'{cnpj};{acordo};{sit_guia};{sit_atrasados}')

    return '\n'.join(textos)


def get_tipo():
    tipos =  {'pgfn': ('pgfn', '36', '92'), 'rfb': ('', '38', '94')}

    text = 'Quais parcelamentos serão gerados?'
    tipo = confirm(title='Tipo de parcelamento', text=text, buttons=('pgfn', 'rfb'))
    if tipo is None: return False

    return tipos[tipo]


@time_execution
def run():
    prev_cert, session = None, None
    empresas = open_lista_dados()
    if not empresas: return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None: return False

    tipo = get_tipo()
    if not tipo: return False

    venc = get_comp(subject='dt vencimento', printable='dd/mm/yyy', strptime='%d/%m/%Y')
    if not venc: return False

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
            text = consulta_parcelamento(cnpj, tipo, venc, session)
        else:
            text = f'Erro Login - {res["value"]}'

        print(cnpj, text)
        escreve_relatorio_csv(f"{cnpj};{text}")

    header = "cnpj;acordo;sit. guia do mes atual;sit. guias em atraso"
    escreve_header_csv(header)

    if isinstance(session, Session): session.close()
    return True


if __name__ == '__main__':
    run()