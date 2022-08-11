# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from datetime import datetime
from pyautogui import prompt
from Dados import empresas
import sys
sys.path.append("..")
from comum import atualiza_info, salvar_arquivo, login_usuario, \
escreve_relatorio_csv, escreve_header_csv


def build_info(soup, info):
	print('>>> Montando dados')

	meses_selecionados = ''
	for i in range(0,12):
		mes = str(i).rjust(2,'0')
		try:
			children = soup.find('table', attrs={'id':f'MainContent_rptContaFiscalMes_gdvResultadoDetalhe_{i}'}).findAll('tr', attrs={'class':'linhaNegrito1'})
			for j in range(0, len(children)):
				linha = str(j+2).rjust(2,"0")
				prefix = f'ctl00$MainContent$rptContaFiscalMes$ctl{mes}$gdvResultadoDetalhe$ctl{linha}'
				info[f'{prefix}$hdfReferencia'] = soup.find('input', attrs={'name':f'{prefix}$hdfReferencia'}).get('value')
				info[f'{prefix}$hdfMicroFilme'] = soup.find('input', attrs={'name':f'{prefix}$hdfMicroFilme'}).get('value')
				info[f'{prefix}$hdfTipoDeIdentificador'] = soup.find('input', attrs={'name':f'{prefix}$hdfTipoDeIdentificador'}).get('value')
				info[f'{prefix}$hdfComplementoIdentificador'] = soup.find('input', attrs={'name':f'{prefix}$hdfComplementoIdentificador'}).get('value')
				info[f'{prefix}$hddCodigoAgencia'] = soup.find('input', attrs={'name':f'{prefix}$hddCodigoAgencia'}).get('value')
				info[f'{prefix}$hddCodigoBanco'] = soup.find('input', attrs={'name':f'{prefix}$hddCodigoBanco'}).get('value')
				info[f'{prefix}$hddDataRecepcao'] = soup.find('input', attrs={'name':f'{prefix}$hddDataRecepcao'}).get('value')
				info[f'{prefix}$hddValorTotalGare'] = soup.find('input', attrs={'name':f'{prefix}$hddValorTotalGare'}).get('value')
			meses_selecionados = meses_selecionados + info[f'{prefix}$hdfReferencia'] + ","
		except:
			pass
	info['ctl00$MainContent$hfMesReferenciaSelecionado'] = meses_selecionados[:-1]
	return info


def consulta_sit_fiscal(cnpj, session, s_id, ano):
	print(f'>>> Consultando situação fiscal')
	url = "https://www10.fazenda.sp.gov.br/ContaFiscalICMS/Pages/ContaFiscal.aspx"

	info = {'SID': s_id}
	response = session.get(url, params=info)
	state, generator, validation = atualiza_info(response)

	info_pesquisa = {
		'__EVENTTARGET': 'ctl00$MainContent$btnConsultar', '__EVENTARGUMENT': '',
		'__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION': validation, 
		'ctl00$MainContent$hdfConteudoPopupDataInicioAtividade': '',
		'ctl00$MainContent$txtHdfIndicesMesesAbertos': '',
		'ctl00$MainContent$hdfHeight': '789', 'ctl00$MainContent$ddlContribuinte': 'CNPJ', 
		'ctl00$MainContent$txtCriterioConsulta': cnpj, 'ctl00$MainContent$ddlReferencia': ano
	}

	response = session.post(url, params=info, data=info_pesquisa)
	state, generator, validation = atualiza_info(response)
	soup = BeautifulSoup(response.content, 'html.parser')

	check = soup.find('span', attrs={'id':'MainContent_lblMensagemDeErro'})
	if check: return f'{cnpj};{check.text.strip()}'
	
	info_pesquisa = build_info(soup, info_pesquisa)
	info_pesquisa['__VIEWSTATEENCRYPTED']= ''
	info_pesquisa['__EVENTTARGET'] = 'ctl00$MainContent$lnkImprimeContaFiscal'

	info_pesquisa['__VIEWSTATE']= state
	info_pesquisa['__EVENTVALIDATION']= validation
	info_pesquisa['__VIEWSTATEGENERATOR']= generator

	print('>>> Salvando relatorio')
	response = session.post(url, params=info, data=info_pesquisa)
	filename = response.headers.get('content-disposition', '')
	if not filename: return f'{cnpj};Não havia arquivo'

	nome = f"{cnpj};INF_FISC_REAL;Situação Fiscal.pdf"
	salvar_arquivo(nome, response)
	return f'{cnpj};Relatorio disponivel'
			

def run():
	comeco = datetime.now()
	print("Execução iniciada as: ", comeco)
	print("Insira o ano da consulta na caixa de texto... ", end='')

	total, lista_empresas = len(empresas), empresas
	ano = prompt("Digite o ano da consulta 0000:")
	if not ano:
		print("Execução abortada")
		return False

	print(ano)
	for i, empresa in enumerate(lista_empresas):
		cnpj = empresa[0]

		print(i, total)
		s, s_id = login_usuario(*empresa)

		if s:
			texto = consulta_sit_fiscal(cnpj, s, s_id, ano)
			s.close()
		else:
			texto = f'{cnpj};{s_id}'

		escreve_relatorio_csv(texto)
		print(f'>>> {texto}\n\n', end='')

	escreve_header_csv('cnpj;situacao')
	print("Tempo de execução: ", datetime.now() - comeco)
	return True


if __name__ == '__main__':
	run()
	