# -*- coding: utf-8 -*-
import time
from bs4 import BeautifulSoup
from dateutil.relativedelta import *
from datetime import datetime, date
from requests import Session
from pyautogui import prompt
from time import sleep
import os, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice

'''
Download dos pdf no período dtInicio-dats_final para a empresa CNPJ (notas prestadas e tomadas)
Os arquivos são obtidos pelo site 'https://valinhos.sigissweb.com' através do usuario e senha 
do contribuinte.

url_acesso -> link para acessar o sistema
url_pesquisa -> link para filtrar o período necessário
url_sintético -> link que retorna o arquivo .pdf sintético do período solicitado
url_analitico -> link que retorna o arquivo .pdf analítico do período solicitado
'''


def consulta_xml(empresa, data_inicio, data_final):
    cnpj, senha, nome = empresa
    nome = nome.replace('/', ' ').replace('   ', ' ').replace('  ', ' ')
    
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    url_pesquisa = "https://valinhos.sigissweb.com/nfecentral?oper=efetuapesquisasimples"
    url_sintetico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfesintetico"
    url_analitico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfeanalitico"
    # obter login e senha do banco de dados utilizando o cnpj

    # inicia a sessão no site, realiza o login e obtém os arquivos para o contribuinte
    with Session() as s:
        login_data = {"loginacesso": cnpj, "senha": senha}
        pagina = s.post(url_acesso, login_data)

        try:
            soup = BeautifulSoup(pagina.content, 'html.parser')
            soup = soup.prettify()
            # print(soup)
            regex = re.compile(r"'Aviso', '(.+)<br>")
            regex2 = re.compile(r"'Aviso', '(.+)\.\.\.', ")
            try:
                documento = regex.search(soup).group(1)
            except:
                documento = regex2.search(soup).group(1)
            _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}')
            print(f"❌ {documento}")
            return False
        except:
            pass
    
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
            'data_emissao': '',
            'operData': '=',
            'data_emissao2': '',
            'mesi': data_inicio[0],
            'anoi': data_inicio[1],
            'mesf': data_final[0],
            'anof': data_final[1],
            'aliq_iss': '',
            'regime': '?',
            'iss_retido': '?',
            'cancelada': '?',
            'tipoPesq': 'normal'
        }

        pagina = s.post(url_pesquisa, info)
        if pagina.status_code != 200:
            _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Erro ao acessar a página')
            print('>>> Erro ao acessar a página.')
            return False
    
        if data_inicio != data_final:
            file = f'{nome} - {cnpj} - Sintético - {str(data_inicio[0])}-{str(data_inicio[1])} até {str(data_final[0])}-{str(data_final[1])}'
        else:
            file = f'{nome} - {cnpj} - Sintético - {str(data_inicio[0])}-{str(data_inicio[1])}'
            
        pasta = 'execução/Sintético'

        salvar = [(s.get(url_sintetico)), (s.get(url_analitico))]
        # salvar relatório sintético depois relatório analítico
        for url in salvar:
            arquivo_nome = re.compile(r'- \d+ - (.+)').search(file).group(1)
            
            # rotina para salvar os arquivos pdf
            if 'text' not in url.headers.get('Content-Type', 'text'):
                os.makedirs('execução/Sintético', exist_ok=True)
                os.makedirs('execução/Analítico', exist_ok=True)
                
                arquivo = open(os.path.join(pasta, file+'.pdf'), 'wb')
                for parte in url.iter_content(100000):
                    arquivo.write(parte)
                arquivo.close()
                
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Relatório {arquivo_nome} gerado')
                print(f"✔ Relatório {arquivo_nome} salvo")
            else:
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Erro ao acessar a página {arquivo_nome}')
                print(f'❌ Não gerou relatório {arquivo_nome}.')

            # necessário para não sobrepor o cachê da pesquisa
            sleep(1)

            if data_inicio != data_final:
                file = f'{nome} - {cnpj} - Analítico - {str(data_inicio[0])}-{str(data_inicio[1])} até {str(data_final[0])}-{str(data_final[1])}'
            else:
                file = f'{nome} - {cnpj} - Analítico - {str(data_inicio[0])}-{str(data_inicio[1])}'
            pasta = 'execução/Analítico'
                
    return True


@_time_execution
def run():
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    data_inicio = prompt("Data inicio no formato 00/0000:").split('/')
    data_final = prompt("Data final no formato 00/0000:").split('/')

    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        try:
            consulta_xml(empresa, data_inicio, data_final)
        except Exception as e:
            print(e)


if __name__ == '__main__':
    run()
