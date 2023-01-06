# -*- coding: utf-8 -*-
import time
from bs4 import BeautifulSoup
from requests import Session
from sys import path
import re

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice, _escreve_header_csv, _download_file
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
                try:
                    response = s.get(_url)
                    print('>>> Abrindo site...')
                    time.sleep(5)
                    # pega o token de verificação do site
                    soup = BeautifulSoup(response.content, 'html.parser')
                    soup = soup.prettify()
                    padrao_response = re.compile(r'__RequestVerificationToken.+value=\"(.+)\"/>')
                    request_verification = padrao_response.search(soup)
                    request_verification = request_verification.group(1)
    
                    # gera o token para passar pelo captcha
                    recaptcha_data = {'sitekey': _site_key, 'url': _url}
                    token = _solve_hcaptcha(recaptcha_data)
    
                    # payload da requisição do site para fazer o login na empresa
                    payload = {'Cnpj': str(cnpj),
                               'h-captcha-response': str(token),
                               '__RequestVerificationToken': str(request_verification)}
                
                    # faz login
                    response = s.post('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes', data=payload)
                    # abre a tela de consulta
                    
                    payload = {'vc': cnpj}
                    s.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes/Home/ConsultarCnpj?vc={}'.format(cnpj), data=payload)
                    # verifica a situação da empresa em relação ao simples nacional e simei
                    print('>>> Verificando situação...')
                    soup = BeautifulSoup(response.content, 'html.parser')
                    soup = soup.prettify()
                    sn = re.compile(r'Situação no Simples Nacional:\n.+\n.......(.+)').search(str(soup))
                    simei = re.compile(r'Situação no SIMEI:\n.+\n.......(.+)').search(str(soup))
                    stylesheet = re.compile(r'\.css\?v=(.+)\" rel=\"stylesheet\"/>').search(str(soup))
                    # print(soup)


                    # abre a tela de mais informações da empresa
                    print('>>> Verificando eventos...')
                    response = s.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes/Home/AjaxMaisInfo?cnpjHdn={}'.format(cnpj))
                    # verifica os eventos futuros em relação ao simples nacional e simei
                    soup = BeautifulSoup(response.content, 'html.parser')
                    soup = soup.prettify()
                    padrao_eventos_sn = re.compile(r'Eventos Futuros \(Simples Nacional\)(\n.+){4}\n...(.+)')
                    padrao_eventos_simei = re.compile(r'Eventos Futuros \(SIMEI\)(\n.+){4}\n...(.+)')

                    # tenta pegar a informação se não conseguir print o código do site e para
                    try:
                        print('>>> Verificando eventos SN...')
                        evento_sn = padrao_eventos_sn.search(str(soup))
                        evento_sn = str(evento_sn.group(2))
                        
                        if evento_sn == '<thead>':
                            padrao_evento_sn = re.compile(r'Data Efeito(\n.+){6}.\n +(.+)\n.+\n.+\n +(.+)')
                            evento = padrao_evento_sn.search(str(soup)).group(2)
                            evento_data = padrao_evento_sn.search(str(soup)).group(3)
                            evento_sn = f'{evento} - {evento_data}'
                            # s = download_pdf(response, empresa)
                    except:
                        # print(soup)
                        exit()

                    try:
                        print('>>> Verificando eventos SIMEI...')
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
                    s.cookies.clear()
                    s.close()

                except:
                    if tentativas >= 5:
                        _escreve_relatorio_csv(';'.join([cnpj, nome, 'Erro no login']), nome='Optante Simples Nacional')
                        print('❌ Erro ao consultar a empresa')
                        s.cookies.clear()
                        s.close()
                        break

                    print('❌ Erro na consulta\n>>> Tentando novamente...')
                    # print(soup)
                    x = 'Erro'
                    tentativas += 1
                    s.cookies.clear()
                    s.close()


def download_pdf(pagina, empresa):
    os.makedirs('execução\documentos', exist_ok=True)
    soup = BeautifulSoup(pagina.content, 'html.parser')
    
    with open('css_situacao_simples.css', 'r') as f:
        css = f.read()
    style = "<style type='text/css'>" + css + "</style>"
    html = BeautifulSoup(f'<html><head><meta charset="UTF-8">{style}</head><body></body></html>', 'html.parser')
    cnpj, nome = empresa
    
    nome = f'{nome} - {cnpj} - Evento Futuro Simples Nacional.pdf'
    with open(nome, 'w+b') as pdf:
        # pisa.showLogging()
        pisa.CreatePDF(str(html), pdf)
    return s



@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    consulta(empresas, index)
    _escreve_header_csv(texto='CNPJ;NOME;OPTANTE DO SIMPLES NACIONAL;OPTANTE DO SIMEI;EVENTOS SN;EVENTOS SIMEI', nome='Optante Simples Nacional.csv')


if __name__ == '__main__':
    run()
