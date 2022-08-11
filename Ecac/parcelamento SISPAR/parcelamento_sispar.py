# -*- coding: utf-8 -*-
from requests.exceptions import ConnectionError
from selenium import webdriver
from bs4 import BeautifulSoup
from pyautogui import prompt
from requests import Session
from urllib import parse
from time import sleep
from sys import path
import json
import os
import sys
sys.path.append("..")
from comum import pfx_to_pem, salvar_arquivo, escreve_relatorio_csv

path.append(r'..\..\_comum')
from ecac_comum import new_session_ecac, new_login_ecac
from comum_comum import time_execution, escreve_relatorio_csv, _escreve_header_csv, open_lista_dados, where_to_start, which_cert, _headers

'''
    #forms ficam no arquivo info_forms.json

    # Verificar 'id' do menu "Emissão de Documento > Documento de Arrecadação" para 
    atualizar a variavel "info_parcelas"

    # Variavel "info_lista" também sofre atualização

    # Verificar 'id' dos botões "Nova Impressão" e "Reimprimir" para atualizar
    campos dos forms "info_escolha_reimprime" e "info_reimprime"
'''


def get_ViewState(pagina, parcelamento=False):
    tipos_parcelamentos = [
        'DEFERIDO E CONSOLIDADO', 'AGUARDANDO PROVIDENCIAS ADMINISTRATIVAS',
        # 'AGUARDANDO PAGAMENTO',
    ]

    soup = BeautifulSoup(pagina.content, 'html.parser')
    ViewState = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
    if not parcelamento: return ViewState

    attrs = {'id': "formListaDarf:grupoCpfCnpj:0:listaParc:idListaParcelamentos_data"}
    table_body = soup.find('tbody', attrs=attrs)
    linhas = table_body.find_all('tr')

    acordos = []
    for linha in linhas:
        colunas = linha.find_all('td', attrs={'class': "colunaAlinhaCentro"})
        if colunas[2].text not in tipos_parcelamentos: continue
        acordos.append(colunas[1].text.strip())

    print('>>> Acordos consolidados:', len(acordos))
    return ViewState, acordos


def get_parcelas(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    try:
        attrs = {'id': "formResumoParcelamentoDarf:tabaplicacaoparcelamentoparcela_data"}
        update = soup.find('update', attrs={'id': "formResumoParcelamentoDarf"})
        
        soup = BeautifulSoup(update.text, 'html.parser')
        table_body = soup.find('tbody', attrs=attrs)
        linhas = table_body.find_all('tr')
    except:
        return []
    
    parcelas = []
    for linha in linhas:
        colunas = linha.find_all('td')

        dados = []
        for i, ele in enumerate(colunas):
            if i not in [0, 3, 6]: continue

            find_ele = ele.find('img')
            if find_ele:
                dados.append(find_ele.get('title'))
            else:
                dados.append(ele.text)
        parcelas.append(dados)
    return parcelas


def verifica_erro(pagina):
    soup = BeautifulSoup(pagina, 'html.parser')
    try:
        div_acesso = soup.find('div', attrs={'id':'perfilAcesso'})
        msg = div_acesso.p.text.strip()
    except:
        if soup.find('div', attrs={'id': 'dialog-bloqueio-ativo-caixapostal'}):
            msg = ''
        else:    
            msg = 'Ecac indisponivel'

    return msg.replace('ATENÇÃO: ', '').replace('ATENÇÃO:', '')


def login(cnpj, certificado, senha):
    print('>>> Logando', cnpj)

    query = {
        'response_type': 'code', 'client_id': 'cav.receita.fazenda.gov.br',
        'scope': 'openid+govbr_recupera_certificadox509+govbr_confiabilidades',
        'redirect_uri': 'https://cav.receita.fazenda.gov.br/autenticacao/login/govbrsso'
    }

    url_base = 'https://certificado.sso.acesso.gov.br/authorize'
    url_aux = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    url_acompanha = 'https://cav.receita.fazenda.gov.br//Servicos/ATBHE/PGFN/acompanharRequerimentos/app.aspx'

    parsed_url = parse.urlparse(url_base)._replace(query=parse.urlencode(query, safe='+'))
    url_login = parse.urlunparse(parsed_url)

    with pfx_to_pem(certificado, senha) as cert:
        driver =  os.path.join('..', 'phantomjs.exe')
        args = ['--ssl-client-certificate-file='+cert]

        driver = webdriver.PhantomJS(driver, service_args=args)
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()
        driver.get(url_aux)
        sleep(1)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    base_state = soup.find('div', attrs={'id':'login-dados-certificado'}).a.get('href')
    state = base_state.split('=')[-1]

    query = {
        'response_type': 'code', 'client_id': 'cav.receita.fazenda.gov.br',
        'scope': 'openid+govbr_recupera_certificadox509+govbr_confiabilidades',
        'redirect_uri': 'https://cav.receita.fazenda.gov.br/autenticacao/login/govbrsso',
        'state': state
    }

    url_base = 'https://certificado.sso.acesso.gov.br/authorize'
    parsed_url = parse.urlparse(url_base)._replace(query=parse.urlencode(query, safe='+'))
    url_login = parse.urlunparse(parsed_url)

    try:
        driver.get(url_login)
        sleep(1)
        driver.execute_script(f"document.getElementById('txtNIPapel2').value = '{cnpj}';")
        driver.execute_script(f"enviaDados('formPJ');")
        sleep(2)

        msg = verifica_erro(driver.page_source)
        if msg:
            return False, msg

        driver.get(url_acompanha)
        sleep(2)
        script_js = "JSON.parse(localStorage.getItem('currentUser')).token;"
        token = driver.execute_script(f"return {script_js}")
    except:
        driver.quit()
        return False, 'Ecac indisponivel - No token'

    cookies = driver.get_cookies()
    driver.quit()

    if not cookies:
        return False, 'Ecac indisponivel - No cookies'

    session = Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    return session, token


def consulta_parc(cnpj, session, venc):
    print('>>> Consultando parcelamento')
    url_login = 'https://www2.pgfn.fazenda.gov.br/ecac/contribuinte/loginJwt.jsf?fn=adesaoParcelamentos&token='

    form_base = 'formResumoParcelamentoDarf'
    url_base = 'https://sisparnet.pgfn.fazenda.gov.br/sisparInternet/autenticacao.jsf'
    url_auth = f'{url_base}/autenticacao.jsf'
    url_darf = f'{url_base}/listaParcelamentosDarf.jsf'
    url_post = f'{url_base}/internet/darf/resumoParcelamentoDarfInternet.jsf'
    url_download = f'{url_base}/internet/darf/resultadoDarfInternet.jsf'

    with open('infos.json') as f:
        infos = json.load(f)

    pagina = session.get(url_base, headers=_headers, verify=False)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    try:
        ID = soup.find('input', attrs={'id': 'paramecac'}).get('value')
    except:
        return f'{cnpj};;;;Falha ao entrar no SISPAR'

    # ACESSA PÁGINA DE CONSULTA DOS PARCELAMENTOS
    info = {'paramecac': ID}
    pagina = session.post(url_auth, info, headers=_headers, verify=False)
    try:
        ViewState, acordos = get_ViewState(pagina, True)
        if not len(acordos): return f'{cnpj};;;;Sem Acordos Consolidados'
    except IndexError:
        return f'{cnpj};;;;Falha ao carregar acordos'
    
    textos = []
    for acordo in acordos:
        print('>>> Consultando acordo', acordo)

        # ACESSO LISTA DARF/DAS
        infos['info_parcelas']['javax.faces.ViewState'] = ViewState
        pagina = session.post(url_darf, infos['info_parcelas'], headers=_headers, verify=False)
        ViewState = get_ViewState(pagina)
        
        # CLICA NO ACORDO
        infos['info_acordo']['formListaDarf:grupoCpfCnpj:0:listaParc:idListaParcelamentos_instantSelectedRowKey'] = int(acordo)
        infos['info_acordo']['formListaDarf:grupoCpfCnpj:0:listaParc:idListaParcelamentos_selection'] = int(acordo)
        infos['info_acordo']['javax.faces.ViewState'] = ViewState
        pagina = session.post(url_darf, infos['info_acordo'], headers=_headers, verify=False)
        sleep(1)

        # ACESSA AS PARCELAS
        infos['info_clica']['formListaDarf:grupoCpfCnpj:0:listaParc:idListaParcelamentos_selection'] = int(acordo)
        infos['info_clica']['formListaDarf:parcelamentoSelecionado'] = int(acordo)
        infos['info_clica']['javax.faces.ViewState'] = ViewState
        pagina = session.post(url_darf, infos['info_clica'], headers=_headers, verify=False)
        pagina = session.get(url_post, headers=_headers, verify=False)
        ViewState = get_ViewState(pagina)
        
        infos['info_lista']['javax.faces.ViewState'] = ViewState
        pagina = session.post(url_post, infos['info_lista'], headers=_headers, verify=False)

        parcelas = get_parcelas(pagina)
        if not parcelas:
            texto = f'{cnpj};{acordo};;;Falha ao carregar parcelas'
            textos.append(texto)
            continue

        for i, parcela in enumerate(parcelas):
            p_self, p_venc, p_sit = parcela
            
            if p_venc != venc:
                if p_sit == 'Vencido':
                    textos.append(f'{cnpj};{acordo};{p_self};{p_venc};{p_sit}')
                continue

            print(f'>>> Parcela {p_self}, {p_sit}')
            form = f'{form_base}:tabaplicacaoparcelamentoparcela:{i}'

            if p_sit == 'Não emitido':
                form += ':NNimpressaoDarf'
                for i in range(5):
                    # ESCOLHE PARCELA
                    infos['info_escolha_imprime']['javax.faces.ViewState'] = ViewState
                    infos['info_escolha_imprime']['javax.faces.source'] = form
                    infos['info_escolha_imprime'][form] = form
                    pagina = session.post(url_post, infos['info_escolha_imprime'])
                    
                    # CALCULA GUIA
                    aux_soup  = BeautifulSoup(pagina.content, 'html.parser')
                    valor = aux_soup.find('input', attrs={'id':'formResumoParcelamentoDarf:valorReceita'})
                    valor = valor.get('value', '') if valor else 0.00
                    infos['info_calcula']['formResumoParcelamentoDarf:valorReceita'] = valor
                    infos['info_calcula']['javax.faces.ViewState'] = ViewState
                    pagina = session.post(url_post, infos['info_calcula'])
                    
                    # GERA O ARQUIVO
                    infos['info_gerar']['javax.faces.ViewState'] = ViewState
                    pagina = session.post(url_post, infos['info_gerar'])

                    # DOWNLOAD DO ARQUIVO
                    infos['info_imprime']['javax.faces.ViewState'] = ViewState
                    salvar = session.post(url_download, infos['info_imprime'])
                    filename = salvar.headers.get('content-disposition', '').lstrip("filename=")
                    if filename: break

            elif p_sit == 'Já emitido':
                form += ':NNreImpressaoDarf'
                for i in range(3):
                    # GERA O ARQUIVO
                    infos['info_escolha_reimprime']['javax.faces.ViewState'] = ViewState
                    infos['info_escolha_reimprime']['javax.faces.source'] = form
                    infos['info_escolha_reimprime'][form] = form
                    pagina = session.post(url_post, infos['info_escolha_reimprime'])

                    # DOWNLOAD DO ARQUIVO
                    infos['info_reimprime']['javax.faces.ViewState'] = ViewState
                    salvar = session.post(url_post, infos['info_reimprime'])
                    filename = salvar.headers.get('content-disposition', '').lstrip("filename=")
                    if filename:
                        break
            else:
                textos.append(f'{cnpj};{acordo};{p_self};{p_venc};{p_sit}')
                continue

            texto = f'{cnpj};{acordo};{p_self};{p_venc}'
            if filename:
                str_venc = venc.replace('/', '-')
                _, mes, ano = venc.split('/')

                nome = (
                    f'{cnpj};IMP_FISCAL;{str_venc};SISPAR DÍVIDA ATIVA;' +
                    f'{mes};{ano};PARCELAMENTO;DAS - {acordo} - p {p_self}.pdf'
                )
                salvar_arquivo(nome, salvar)
                texto += ';Guia disponivel'
            else:
                texto += ';Erro de download'

            textos.append(texto)

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
                text = consulta_parc(cnpj, session, venc)
            except ConnectionError:
                text = f'{cnpj};Erro Login - {res["value"]}'
                print('\t--connection error')
        else:
            text = f'{cnpj};Erro Login - {res["value"]}'
            print(text)

        escreve_relatorio_csv(text)

    if session is not None:
        session.close()

    header = "cnpj;acordo;sit. guia do mes atual;sit. guias em atraso"
    _escreve_header_csv(header)


if __name__ == '__main__':
    run()
