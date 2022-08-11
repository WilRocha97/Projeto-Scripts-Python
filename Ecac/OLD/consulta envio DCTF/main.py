# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from Dados import empresas
import sys
sys.path.append("..")
from comum import new_session, new_login, escreve_relatorio_csv, \
escreve_header_csv, time_execution


def consulta_envio_dctf(cnpj, comp, s):
    headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

    url_home = 'https://cav.receita.fazenda.gov.br/ecac'
    url = "https://cav.receita.fazenda.gov.br/Servicos/ATSPO/DCTF/Consulta/Abrir.asp"

    response = s.get(url_home, headers=headers, verify=False)

    print("Verificando Tabela")
    response = s.get(url, headers=headers)
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', attrs={'id':'tbDeclaracoes'})

    if not table: return False
    linhas = table.findAll('tr')

    if len(linhas) < 2:
        print(">>>Não possui declaração")
        escreve_relatorio_csv(f"{cnpj};Não possui declaração")
        return True

    for linha in linhas:
        valores = linha.findAll('td')
        if not valores: continue
        if valores[1].text.lower() != comp.lower(): continue

        print(">>>Possui declaração no periodo")
        escreve_relatorio_csv(f"{cnpj};" + ";".join(td.text for td in valores[1:-1]))
        break
    else:
        print(">>>Não possui declaração no periodo")
        escreve_relatorio_csv(f"{cnpj};Não possui declaração no periodo {comp}")

    return True


@time_execution
def run():
    prev_cert, s = [None] * 2
    lista_empresas = empresas
    total = len(lista_empresas)

    comp = input('competencia como no exemplo - Janeiro/2020:')
    for i, empresa in enumerate(lista_empresas):
        cnpj, *cert = empresa

        if cert != prev_cert:
            if s: s.close()
            s = new_session(*cert)
            prev_cert = cert
            if not s:
                prev_cert = None
                continue

        print(i, total)
        res = new_login(cnpj, s)
        if res['Key']:
            if not consulta_envio_dctf(cnpj, comp, s):
                print('Empresa não carregou a tabela...')
        else:
            print('>>>Não logou')
            escreve_relatorio_csv(f'{cnpj};{res["Value"]}')

        print('\n', end='')
    else:
        if s != None: s.close()

    escreve_header_csv("cnpj;periodo;dt. recepcao;periodo inicial;periodo final;situacao;tipo/status")


if __name__ == '__main__':
    run()