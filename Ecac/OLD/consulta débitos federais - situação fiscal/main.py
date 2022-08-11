# -*- coding: utf-8 -*-
from datetime import datetime
from requests import Session
from bs4 import BeautifulSoup
from Dados import empresas
from time import sleep
import pdfplumber
import os, re, sys
sys.path.append("..")
from comum import new_login, new_session, salvar_arquivo, \
escreve_relatorio_csv, time_execution


def pegar_texto(arq):
    texto = ''
    caminho = os.path.join('documentos', arq)
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages[:]:
            texto_pagina = pagina.extract_text()
            if not texto_pagina: continue
            texto += texto_pagina + '\n'

    return texto


def analisa_ocorrencias(razao, cnpj, arquivo):
    print('>>>Verificando ocorrencias')
    re_pendencia = re.compile(r'^(\w{5,} .+) _{26,}$')
    str_rec = '_'*37+' Diagnóstico Fiscal na Receita Federal '+'_'*37
    str_pen_rec = 'Não foram detectadas pendências/exigibilidades suspensas para o contribuinte nos controles da Receita Federal.'
    str_faz = '_'*26+' Diagnóstico Fiscal na Procuradoria-Geral da Fazenda Nacional '+'_'*25
    str_pen_faz = 'Não foram detectadas pendências/exigibilidades suspensas para esse contribuinte nos controles da Procuradoria-Geral da Fazenda Nacional.'
    str_rec_faz = '_'*16+' Diagnóstico Fiscal na Receita Federal e Procuradoria-Geral da Fazenda Nacional '+'_'*15
    str_pen_rec_faz = 'Não foram detectadas pendências/exigibilidades suspensas nos controles da Receita Federal e da Procuradoria-Geral da Fazenda Nacional.'

    linhas = pegar_texto(arquivo).split('\n')
    gravar = False
    for linha in linhas:
        if gravar:
            if linha in [str_pen_faz, str_pen_rec, str_pen_rec_faz]:
                escreve_relatorio_csv(f'{razao};{cnpj};{gravar};{linha}')
                gravar = False
                continue
            ocorrencia = re_pendencia.search(linha)
            if ocorrencia:
                escreve_relatorio_csv(f'{razao};{cnpj};{gravar};{ocorrencia.group(1)}')
        if linha == str_rec_faz:
            gravar = 'Receita Federal e Procuradoria-Geral da Fazenda Nacional'
        if linha == str_rec:
            gravar = 'Receita Federal'
        if linha == str_faz:
            gravar = 'Procuradoria-Geral da Fazenda Nacional'

    print('>>>Verificacao terminada.')


def consulta_debitos(cnpj, s):
    print('>>>Consultando débitos')
    headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

    razao_regex = re.compile(r'- (.*)')
    cnpj_regex = re.compile(r'Titular \(Acesso GOV\.BR por Certificado\): ([\d\.\/\-]{14,18})')
    
    url_home = 'https://cav.receita.fazenda.gov.br/ecac'
    url_iduser = 'https://www2.cav.receita.fazenda.gov.br/Servicos/ATSPO/eSitFis.app/IdentificaUsuario/index.asp'
    url_nova_consulta = 'https://www2.cav.receita.fazenda.gov.br/Servicos/ATSPO/eSitFis.app/GerenciaPedido/PedidoConsultaFiscal.asp'
    url_salvar = 'https://www2.cav.receita.fazenda.gov.br/Servicos/ATSPO/IntegraSitfis/RelatorioEcac.aspx'

    pagina = s.get(url_home, headers=headers)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    titular, razao = soup.find('div', attrs={'id':'informacao-perfil'}).find_all('span')

    titular = cnpj_regex.search(titular.text)
    titular = ''.join(i for i in titular.group(1) if i.isdigit())
    razao = razao_regex.search(razao.text)
    razao = razao.group(1).strip()

    #identificar o user
    response = s.post(url_iduser, headers=headers, verify=False)

    #gerar nova consulta
    info_nova_consulta = {'IndNovaConsulta':'true','OpConsulta':'1'}
    response = s.get(url_nova_consulta, params=info_nova_consulta, headers=headers)

    string_cookie = ''
    for cookie in s.cookies.get_dict():
        string_cookie = string_cookie + f'{cookie}={s.cookies.get_dict()[cookie]};'

    #salvar relatorio
    param_salva ={'TipoNIPesquisa': '1', 'NIPesquisa': str(cnpj), 'Ambiente':'2', 'NICertificado': titular,
    'TipoNICertificado':'1','SistemaChamador':'0101'}

    headers={
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br', 'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7', 'Connection': 'keep-alive',
        'Cookie': string_cookie[:-1], 'Host': 'www2.cav.receita.fazenda.gov.br',
        'Referer': 'https://www2.cav.receita.fazenda.gov.br/Servicos/ATSPO/eSitFis.app/Relatorio/GerarRelatorio.asp',
        'Sec-Fetch-Dest': 'iframe', 'Sec-Fetch-Mode': 'navigate', 'Sec-Fetch-Site': 'same-origin', 'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36'
    }

    while True:
        response = s.get(url_salvar, params=param_salva, headers=headers)
        filename = response.headers.get('content-disposition', '')
        size = response.headers.get('Content-Length', 0)
        if filename:
            if int(size) < 10000:
                sleep(1)
                print('  Falha no Download:', size)
                continue
            else:
                nome =  "".join(cnpj) + ";INF_FISC_REAL;Debitos Federais.pdf"
                salvar_arquivo(nome, response)
                analisa_ocorrencias(razao, cnpj, nome)
                break
        else:
            break

    return True


@time_execution
def run():
    prev_cert, s = [None] * 2
    lista_empresas = empresas
    total = len(lista_empresas)
    escreve_relatorio_csv('Razao;CNPJ;Diagnostico;Pendencia')

    for i, empresa in enumerate(lista_empresas, 1):
        cnpj, *cert = empresa

        if cert != prev_cert:
            if s: s.close()
            s = new_session(*cert)
            prev_cert = cert
            if not s:
                prev_cert = None
                continue

        res = new_login(cnpj, s)
        if res['Key']:
            consulta_debitos(cnpj, s)
        else:
            print('>>>Não logou')
            escreve_relatorio_csv(f'?;{cnpj};{res["Value"]}')

        print(f'<{i}/{total} concluidos>\n')
    else:
        if s != None: s.close()
        

if __name__ == '__main__':
    run()