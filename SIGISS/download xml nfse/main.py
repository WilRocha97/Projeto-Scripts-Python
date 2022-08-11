# -*- coding: utf-8 -*-
from pyautogui import prompt, confirm
from dateutil.relativedelta import relativedelta
from datetime import datetime, date
from requests import Session
from functools import wraps
from Dados import empresas
from time import sleep
import os

# decorator
def time_execution(func):
    @wraps(func)
    def wrapper():
        comeco = datetime.now()
        print("Execução iniciada as:", comeco)
        func()
        print("Tempo de execução:", datetime.now() - comeco)

    return wrapper


def list_comps(dti, dtf):
    dti, dtf = dti.split('/'), dtf.split('/')
    dti = date(int(dti[1]), int(dti[0]), 1)
    dtf = date(int(dtf[1]), int(dtf[0]), 1)

    comps = []
    while True:
        tupla_comp = (str(dti.month).rjust(2, '0'), str(dti.year))
        comps.append(tupla_comp)

        if dti == dtf: break
        dti = dti + relativedelta(months=+1)

    return tuple(comps)


def consulta_xml(cnpj, pwd, comps, tipos):
    print('>>> consultando xml')
    s = Session()
    os.makedirs('docs', exist_ok=True)

    url_acesso = 'https://valinhos.sigissweb.com/ControleDeAcesso'
    url_pesquisa = 'https://valinhos.sigissweb.com/lancamentocentral?oper=efetuapesquisasimples'
    url_download = 'https://valinhos.sigissweb.com/lancamentocentral?oper=geraxml'
    
    infos = {
        'tomador_prestador': '', 'cnpj_cpf_decl': '', 'selcnpjcpfoperdecl': 'EX',
        'cnpj_cpf_dest': '', 'selcnpjcpfoperdest': 'EX', 'cidade_dest': '',
        'selcidadedestoperdest': 'EX', 'mesi': '', 'anoi': '', 'mesf': '', 'anof': '', 
        'mesicaixa': '', 'anoicaixa': '', 'mesfcaixa': '', 'anofcaixa': '', 'iss_retido_fonte': '?', 
        'regime': '?', 'movimento': '?', 'documento_fiscal_canc': '?', 'num_docu_fiscal_pesq': '',
        'serie_docu_fiscal': '', 'data': '', 'aliquota': '?', 'caixaindicado': '?', 'classif':'?'
    }

    login_data = { "loginacesso": cnpj, "senha": pwd }
    response = s.post(url_acesso, login_data)
    if response.status_code != 200: return 'Erro a acessar página'

    msgs = []
    for mes, ano in comps:
        for tipo in tipos:
            infos['tomador_prestador'] = tipo
            infos['mesi'], infos['anoi'] = str(mes), ano
            infos['mesf'], infos['anof'] = str(mes), ano

            response = s.post(url_pesquisa, infos)
            if response.status_code != 200:
                msgs.append(f'{cnpj};{tipo};{mes}-{ano};Erro a acessar')
                continue

            response = s.get(url_download)
            if 'text' in response.headers.get('Content-Type', 'text'):
                msgs.append(f'{cnpj};{tipo};{mes}-{ano};Não existia arquivo disponivel')
                continue

            nome = f'{cnpj} - {str(mes)}-{ano}{tipo}.docs'
            with open(os.path.join('docs', nome), 'wb') as f:
                for bite in response.iter_content(100000):
                    f.write(bite)

            sleep(1)
            msgs.append(f'{cnpj};{tipo};{mes}-{ano};Arquivo xml baixado')

    s.close()
    return msgs


@time_execution
def run():
    dti = prompt("Mes de inicio da consulta (00/0000):")
    dtf = prompt("Mes de final da consulta (00/0000):")
    comps = list_comps(dti, dtf)
    if not comps:
        print("Intervalo de datas invalido")
        return False

    dtipos = {'Prestador': ('P',), 'Tomador': ('T',), 'Ambos': ('P', 'T')}
    tipos = confirm(text="Tipo de consulta", buttons=["Prestador", "Tomador", "Ambos"])
    if not tipos:
        print("Execução abortada")
        return False
    print(f'Consultando tipos {tipos} de {dti} a {dtf}\n')

    tipos = dtipos[tipos]
    total, lista_empresas = len(empresas), empresas
    for i, dados in enumerate(lista_empresas):
        print(i, total)

        msgs = consulta_xml(*dados, comps, tipos)
        print(f'>>>', '\n'.join(msgs), '\n')

    return True


if __name__ == '__main__':
    run()
