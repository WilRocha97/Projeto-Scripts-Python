# -*- coding: utf-8 -*-
from requests.exceptions import Timeout
from bs4 import BeautifulSoup
from pyautogui import prompt
from Dados import empresas
from urllib import parse
import sys
sys.path.append("..")
from comum import time_execution, new_login_credential, escreve_header_csv, \
escreve_relatorio_csv


def extrair_info(soup, cod):
	tipo = {
		'100': 'tributaveis',
		'200': 'sujeitos a tributação com exigibilidade suspensa',
		'300': 'sujeitos a tributacao exclusiva',
		'500': 'recebidos acumuladamente',
		'600': 'isentos - sociedade em conta de participacao'
	}

	tables = soup.find('div', attrs={'id': f'D_RendC{cod}_LINHAS'})
	linhas = tables.find_all('tr', attrs={'id': ''})
	quadros_multilinha = ['100', '200', '300', '500']

	aux, is_new = [], True
	for i, linha in enumerate(linhas):
		infos = [td.text.strip() for td in linha.find_all('td') if td.text.strip()]
		if not infos: continue

		if is_new:
			infos.insert(0, tipo[cod])
			aux.append(infos)
		else:
			infos = infos[1:]
			aux[-1].extend(infos)

		is_new = not is_new if cod in quadros_multilinha else is_new
	
	return aux


def consulta_rend(ni, ano, session):
	print('>>> Consultando rendimentos')

	headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

	texto = []
	quadros = ['100', '200', '300', '500', '600', '700']

	url_base = 'https://cav.receita.fazenda.gov.br'
	url_home = f'{url_base}/ecac/'
	url_rend = f'{url_base}/Servicos/ATCTA/Rendimento/Rend/RendC/RendC.asp'

	response = session.get(url_home, headers=headers)
	soup = BeautifulSoup(response.content, "html.parser")
	razao = soup.find('div', attrs={'id':'informacao-perfil'}).text
	razao = razao.split('-')[-1].strip()

	query = {
		'ano': ano, 'quadro': '', 'ni': ni,
		'tipo_ni': '1', 'nome_ni': razao
	}

	for cod in quadros:	
		query['quadro'] = cod
		url = url_rend.replace('RendC', f'RendC{cod}')

		parsed_url = parse.urlparse(url)._replace(query=parse.urlencode(query))
		url = parse.urlunparse(parsed_url)
		try:
			response = session.post(url, headers=headers, timeout=20)
		except Timeout:
			return f'{razao};{ni};Erro de timeout'

		soup = BeautifulSoup(response.content, 'html.parser')
		header = soup.find('div', attrs={'id': f'D_RendC{cod}_CABEC'}).text
		if not header: continue

		aux = [[razao, ni] + linha for linha in extrair_info(soup, cod)]
		texto.append(aux)

	if not texto:
		return f'{razao};{ni};Nenhum rendimento informado'

	return '\n'.join(';'.join(linha) for quadro in texto for linha in quadro)


@time_execution
def run():
	total, lista_empresas = len(empresas), empresas

	print("Insira o ano da consulta na caixa de texto... ", end='')
	ano = prompt("Digite o ano da consulta:")
	if not ano:
		print("Execução abortada")
		return False

	print(ano)
	for i, empresa in enumerate(lista_empresas):
		razao, ni, *creds = empresa

		print(i, total)
		status, session = new_login_credential(ni, *creds)

		if status:
			texto = consulta_rend(ni, ano, session)
		else:
			texto = f'{razao.strip()};{ni};{session}'

		print(f'>>> {texto}\n\n', end='')
		escreve_relatorio_csv(texto)

	header = (
		'razao;cpf/cnpj;situacao/tipo rendi.;cnpj/cpf fonte pagadora;Fonte pagadora;' +
		'dt. Entrega;rendimento;prev. Oficial;dependentes;pensao alim.;prev. Compl.;' +
		'total deducoes;imposto retido;rendimento isento/sem retencao'
	)
	escreve_header_csv(header)


if __name__ == '__main__':
	run()