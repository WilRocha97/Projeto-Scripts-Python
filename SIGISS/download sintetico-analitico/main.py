# -*- coding: utf-8 -*-
# ! python3
from dateutil.relativedelta import *
from datetime import datetime, date
from requests import Session
from pyautogui import prompt
from Dados import empresas
from time import sleep
import os

'''
Download dos pdf no período dtInicio-dtFinal para a empresa CNPJ (notas prestadas e tomadas)
Os arquivos são obtidos pelo site 'https://valinhos.sigissweb.com' através do usuario e senha 
do contribuinte.

url_acesso -> link para acessar o sistema
url_pesquisa -> link para filtrar o período necessário
url_sintetico -> link que retorna o arquivo .pdf sintetico do período solicitado
url_analitico -> link que retorna o arquivo .pdf analitico do período solicitado
'''


def intervalo_comp(dti, dtf):
    dti = date(int(dti[1]), int(dti[0]), 1)
    dtf = date(int(dtf[1]), int(dtf[0]), 1)

    intervalo = []
    while True:
        tupla_comp = (str(dti.month).rjust(2, '0'), str(dti.year))
        intervalo.append(tupla_comp)
        
        if dti != dtf:
            dti = dti + relativedelta(months=+1)
        else:
            break

    return intervalo


def consulta_xml(nome, cnpj, senha, dtInicio, dtFinal):
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    url_pesquisa = "https://valinhos.sigissweb.com/nfecentral?oper=efetuapesquisasimples"
    url_sintetico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfesintetico"
    url_analitico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfeanalitico"
    # obter login e senha do banco de dados utilizando o cnpj

    # inicia a sessão no site, realiza o login e obtem os arquivos para o contribuinte
    with Session() as s:
        login_data = {"loginacesso":cnpj,"senha":senha}
        pagina = s.post(url_acesso, login_data)
        if pagina.status_code != 200:
            print('>>> Erro a acessar página.')
            return False

        for comp in intervalo_comp(dtInicio, dtFinal):
            info = {
                'cnpj_cpf_destinatario': '',
                'operCNPJCPFdest': 'EX',
                'RAZAO_SOCIAL_DESTINATARIO': '',
                'selnomedestoper': 'EX',
                'id_codigo_servico': '',
                'serie': '',
                'numero_nf': '',
                'operNFE': '=',
                'numero_nf2': '',
                'rps': '',
                'operRPS': '=',
                'rps2': '',
                'data_emissao':'', 
                'operData': '=',
                'data_emissao2': '',
                'mesi': comp[0],
                'anoi': comp[1],
                'mesf': comp[0],
                'anof': comp[1],
                'aliq_iss': '',
                'regime': '?',
                'iss_retido': '?',
                'cancelada': '?',
                'tipoPesq': 'normal'
            }

            file = '-'.join([nome, cnpj, 'Sintetico'])
            pagina = s.post(url_pesquisa, info)
            if pagina.status_code != 200:
                print('>>> Erro a acessar página.')
                return False

            salvar = []          
            salvar.append(s.get(url_sintetico))
            salvar.append(s.get(url_analitico))

            pasta = 'sintetico'

            #salvar relatório sintético depois relatório analítico
            for url in salvar:
                # rotina para salvar os arquivos pdf
                if 'text' not in url.headers.get('Content-Type', 'text'):
                    arquivo = open(os.path.join(pasta, file+'.pdf'), 'wb')
                    for parte in url.iter_content(100000):
                        arquivo.write(parte)
                    arquivo.close()
                    print(f"Arquivo {file}.pdf salvo")

                # necessário para não sobrepor o cache da pesquisa
                sleep(1)

                file = '-'.join([nome, cnpj, 'Analitico'])
                pasta = 'analitico'
                
    return True


if __name__ == '__main__':
    comeco = datetime.now()
    os.makedirs('sintetico', exist_ok=True)
    os.makedirs('analitico', exist_ok=True)

    dtInicio = prompt("Data inicio no formato 00-0000:").split('-')
    dtFinal = prompt("Data inicio no formato 00-0000:").split('-')

    for empresa in empresas:
        print('>>> Buscando empresa', empresa[1])
        try:
            consulta_xml(*empresa, dtInicio, dtFinal)
        except Exception as e:
            print(e)
    fim = datetime.now()
    print("\n>>> Tempo total:", (fim-comeco))
