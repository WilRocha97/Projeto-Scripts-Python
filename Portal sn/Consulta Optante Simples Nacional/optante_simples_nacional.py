# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from sys import path
import re

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice, _escreve_header_csv
from captcha_comum import _solve_hcaptcha

# site key é usado para quebrar o captcha
_site_key = '96a07261-d2ef-4959-a17b-6ae56f256b3f'
_url = 'https://consopt.www8.receita.fazenda.gov.br/consultaoptantes'


def consulta(empresas, index):
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)

        cnpj, nome = empresa

        tentativas = 1
        x = 'Erro'
        while x == 'Erro':

            with Session() as s:
                s.get('https://www.serpro.gov.br/')
                response = s.get(_url)
                print('>>> Abrindo site...')

                # pega o token de verificação do site
                soup = BeautifulSoup(response.content, 'html.parser')
                soup = soup.prettify()
                padrao_response = re.compile(r'__RequestVerificationToken.+value=\"(.+)\"\/>')
                request_verification = padrao_response.search(soup)
                request_verification = request_verification.group(1)

                # gera o token para passar pelo captcha
                recaptcha_data = {'sitekey': _site_key, 'url': _url}
                token = _solve_hcaptcha(recaptcha_data)

                # payload da requisição do site para fazer o login na empresa
                payload = {'Cnpj': str(cnpj),
                           'h-captcha-response': str(token),
                           '__RequestVerificationToken': str(request_verification)}

                try:
                    # faz login
                    s.post('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes', data=payload)
                    # abre a tela de consulta
                    response = s.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes/Home/ConsultarCnpj?vc={}'.format(cnpj))
                    # verifica a situação da empresa em relação ao simples nacional e simei
                    soup = BeautifulSoup(response.content, 'html.parser')
                    soup = soup.prettify()
                    padrao_sn = re.compile(r'Situação no Simples Nacional:\n.+\n.......(.+)')
                    padrao_simei = re.compile(r'Situação no SIMEI:\n.+\n.......(.+)')

                    sn = padrao_sn.search(str(soup))
                    simei = padrao_simei.search(str(soup))

                    # abre a tela de mais informações da empresa
                    response = s.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes/Home/AjaxMaisInfo?cnpjHdn={}'.format(cnpj))
                    # verifica os eventos futuros em relação ao simples nacional e simei
                    soup = BeautifulSoup(response.content, 'html.parser')
                    soup = soup.prettify()
                    padrao_eventos_sn = re.compile(r'Eventos Futuros \(Simples Nacional\)(\n.+){4}\n...(.+)')
                    padrao_eventos_simei = re.compile(r'Eventos Futuros \(SIMEI\)(\n.+){4}\n...(.+)')

                    # tenta pegar a informação se não conseguir print o código do site e para
                    try:
                        evento_sn = padrao_eventos_sn.search(str(soup))
                        evento_sn = str(evento_sn.group(2))
                    except:
                        # print(soup)
                        exit()

                    try:
                        evento_simei = padrao_eventos_simei.search(str(soup))
                        evento_simei = str(evento_simei.group(2))
                    except:
                        # print(soup)
                        exit()

                    sn = str(sn.group(1))
                    simei = str(simei.group(1))

                    print('✔ Consulta realizada')
                    _escreve_relatorio_csv(';'.join([cnpj, nome, sn, simei, evento_sn, evento_simei]), nome='Optante Simples Nacional')
                    x = 'OK'
                    s.close()

                except:
                    '''if tentativas >= 3:
                        _escreve_relatorio_csv(';'.join([cnpj, nome, 'Erro no login']), nome='Optante Simples Nacional')
                        print('❌ Erro ao consultar a empresa')
                        s.close()
                        break'''

                    print('❌ Erro na consulta\n>>> Tentando novamente...')
                    # print(soup)
                    x = 'Erro'
                    tentativas += 1
                    s.close()


@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    consulta(empresas, index)
    cabecalho = ';'.join(['CNPJ', 'NOME', 'OPTANTE DO SIMPLES NACIONAL', 'OPTANTE DO SIMEI', 'EVENTOS SIMPLES NACIONAL', 'EVENTOS SIMEI'])
    _escreve_header_csv(texto=cabecalho, nome='Optante Simples Nacional.csv')


if __name__ == '__main__':
    run()
