# -*- coding: utf-8 -*-
from datetime import datetime
from bs4 import BeautifulSoup
from requests import get
from Dados import empresas
from sys import path
import sys

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice


def consulta_obrigados_sped_fiscal(cnpj):
    url = f'https://www.fazenda.sp.gov.br/sped/obrigados/obrigados.asp'
    
    res = get(f'{url}?CNPJLimpo={cnpj}&Submit=Enviar')
    soup = BeautifulSoup(res.content, 'html.parser')
    linhas = soup.findAll('tr', attrs={'class': 'tabulacao7'})

    if not linhas:
        if 'N&Atilde;O H&Aacute; REGISTROS PARA O CNPJ' in res.text:
            return 'Não há registros a serem consultados'
        else:
            return 'Erro'

    for linha in linhas:
        if not linha.findAll('td')[3].text:
            return 'Existe obrigatoriedade no sped fiscal'
    
    return 'Não possui obrigatoriedade no sped fiscal'


@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)

        text = consulta_obrigados_sped_fiscal(empresa)
        _escreve_relatorio_csv(f'{empresa};{text}')

    print("Tempo de execução: ", datetime.now() - comeco)


if __name__ == '__main__':
    run()
