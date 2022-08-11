# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from Dados import empresas
from pyautogui import confirm
import sys, re, os
sys.path.append("..")
from comum import new_session, new_login, atualiza_info, \
escreve_relatorio_csv, time_execution

_ALLOWED_COLUMNS = [5, 6, 7, 8, 10]

def get_html_doc(cod, cnpj, text):
    # armazena o html da consulta para conferencia
    arq_html = os.path.join('logs', f'{cod} - {cnpj}.html')
    with open(arq_html, 'w') as f:
        f.write(text)


def consulta_pag_darf(cnpj, s, cod_receita, dt_pesquisa):
    print(f'>>>Consultando pagamento de darf', cod_receita)
    url_base = 'https://cav.receita.fazenda.gov.br'
    url_pesquisa = f'{url_base}/Servicos/ATFLA/PagtoWeb.app/Default.aspx'
    url_tabela = f'{url_base}/Servicos/ATFLA/PagtoWeb.app/PagtoWebList.aspx'

    headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

    #voltar a pagina inicial a cada consulta
    s.get(f'{url_base}/ecac', headers=headers)

    response = s.get(url_pesquisa, headers=headers)
    viewstate, viewgenerator, eventvalidation, *aux = atualiza_info(response)
    soup = BeautifulSoup(response.content, 'html.parser')
    text_header = soup.find('span', attrs={'id':'LabelContribuinte'}).text
    razao = re.search(r'Nome:(.+)', text_header).group(1).strip()
    
    dt_inicio, dt_final = dt_pesquisa.split(' - ')
    info = {
        '__EVENTTARGET': '', '__EVENTARGUMENT': '', '__LASTFOCUS': '',
        '__VIEWSTATE': viewstate, '__VIEWSTATEGENERATOR': viewgenerator,
        '__EVENTVALIDATION': eventvalidation, 'campoTipoDocumento': 'Todos',
        'campoDataArrecadacaoInicial': dt_inicio, 'campoDataArrecadacaoFinal': dt_final,
        'campoCodReceita': cod_receita, 'campoNumeroDocumento': '', 'campoValorInicial': '',
        'campoValorFinal': '', 'botaoConsultar': 'Consultar',
    }

    if cod_receita == '': cod_receita = 'normal'
    text_base = f'{razao};{cnpj};{cod_receita};{dt_pesquisa}'

    response = s.post(url_pesquisa, info, headers=headers)
    if 'Nenhuma arrecadação que atenda a estes parâmetros foi localizada.' in response.text:
        print('>>>Nenhuma Arrecadação')
        get_html_doc(cod_receita, cnpj, response.text)
        escreve_relatorio_csv(f'{text_base};Nenhuma Arrecadação')
        return True

    response = s.get(url_tabela, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')
    linhas = soup.find('table', attrs={'id':'listagemDARF'}).find_all('tr')
    for linha in linhas[1:]:
        dados = [campo.text.strip() for i, campo in enumerate(linha.find_all('td')) if i in _ALLOWED_COLUMNS]
        dados.insert(2, dados.pop(0))
        get_html_doc(cod_receita, cnpj, response.text)
        escreve_relatorio_csv(f'{text_base};' + ';'.join(dados))
    print('>>>Dados Salvos')

    return True


@time_execution
def run():

    prev_cert, s = [None] * 2
    lista_empresas = empresas
    os.makedirs('logs', exist_ok=True)
    campos = [
        'Razao' , 'cnpj', 'Tipo pesquisa', 'Periodo pesquisa', 'Dt arrecadacao',
        'Dt venc', 'Dt apuracao', 'Cod receita', 'Valor', 'Valor original',
        'Multa', 'Juros'
    ]
    escreve_relatorio_csv(';'.join(campos))

    dt_normal = input('>>>Periodo da Pesquisa normal 00/00/0000 - 00/00/0000:')
    servs = [('', dt_normal)]
    if confirm('Consulta Trimestral', buttons=["sim", "Não"]) == 'sim':
        dt_trimestral = input('>>>Periodo da Pesquisa trimestral 00/00/0000 - 00/00/0000:')
        servs.extend([('2089', dt_trimestral), ('2372', dt_trimestral)])

    for empresa in lista_empresas:
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
            for serv in servs:
                consulta_pag_darf(cnpj, s, *serv)
        else:
            print('>>>Não logou')
            escreve_relatorio_csv(f'?;{cnpj};{res["Value"]}')

        print('\n', end='')
    else:
        if s != None: s.close()
        

if __name__ == '__main__':
    run()