# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from time import sleep
from sys import path

path.append(r'..\..\_comum')
from captcha_comum import break_normal_captcha
from selenium_comum import exec_phantomjs, get_elem_img
from comum_comum import time_execution, content_or_soup, open_lista_dados, \
where_to_start, escreve_relatorio_csv


# variaveis globais
_data_base =  {"__LASTFOCUS": "", "__EVENTARGUMENT": "", "servidorcaptcha": "P"}

_url = 'https://www.receita.fazenda.gov.br/aplicacoes/ssl/atbhe/codacesso.app/PFCodAcesso.aspx'
_psw = 'Veiga123'


@content_or_soup
def get_info_post(soup):
	infos = [
		soup.find('input', attrs={'id':'__VIEWSTATE'}),
		soup.find('input', attrs={'id':'__VIEWSTATEGENERATOR'}),
		soup.find('input', attrs={'id':'__EVENTVALIDATION'}),
	]

	# state, generator, validation
	return tuple(info.get('value', '') for info in infos if info)


@content_or_soup
def get_info_recibos(soup, recibos):
	info, missing = {}, [] 

	for i in range(2):
		rec = soup.find('span', id='lblReciboAno'+str(i+1))
		rec = rec.text.strip() if rec else ''

		if rec != '':
			if recibos[i] == '':
				missing.append(rec)
				continue

			info.update({'txtReciboAno'+str(i+1):recibos[i]})

	return missing if missing else info


def new_session_senha(cpf, dt_nasc, driver):
	print('\n>>> logando para', cpf, dt_nasc)
	driver.get(_url)
	sleep(1)

	captcha = driver.find_element_by_id('img_captcha_serpro_gov_br')
	if not captcha : return 'Erro Login - não encontrou captcha'

	captcha = break_normal_captcha(get_elem_img(captcha, driver))
	if isinstance(captcha, str): return captcha

	state, generator, validation = get_info_post(content=driver.page_source)
	cpf = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:])

	data = _data_base.copy()
	data.update({
		'__EVENTTARGET': 'btnAvancar', '__VIEWSTATE': state,
		'__EVENTVALIDATION': validation, '__VIEWSTATEGENERATOR': generator,
		'txtCPF': cpf, 'txtDtNascimento': dt_nasc, 'txtCaptcha': captcha['code']
	})

	session = Session()
	for cookie in driver.get_cookies():
		session.cookies.set(cookie['name'], cookie['value'])

	res = session.post(_url, data, verify=False)
	return res, session


def gera_senha_receita(recibos, res, session):
	print('>>> alterando senha para', _psw)

	soup = BeautifulSoup(res.content, 'html.parser')

	state, generator, validation = get_info_post(soup=soup)
	recibos = get_info_recibos(recibos, soup=soup)
	if isinstance(recibos, list):
		return 'Recibo faltantes - ' + ''.join(recibos)

	data = _data_base.copy()
	data.update(recibos)
	data.update({
		'__EVENTTARGET': 'btnGerarCodigo', '__VIEWSTATE': state,
		'__EVENTVALIDATION': validation, '__VIEWSTATEGENERATOR': generator,
		'txtSenha':_psw, 'txtConfirmaSenha': _psw
	})

	res = session.post(_url, data, verify=False)
	soup = BeautifulSoup(res.content, 'html.parser')

	erro = soup.find('div', id='ValidationSummary2')
	if erro:
		return erro.text.strip()
	else:
		cod = soup.find('span', id='lblCodAcessoPnl')
		val = soup.find('span', id='lblpnlValidade')
		if all((cod, val)):
			return f'{cod.text.strip()};{val.text.strip()}'

	raise Exception(soup.prettify())


@time_execution
def run():
	empresas = open_lista_dados()
	if not empresas: return False

	index = where_to_start(tuple(i[0] for i in empresas))
	if index is None: return False

	driver = exec_phantomjs()
	if isinstance(driver, str):
		raise Exception(driver)

	for cpf, dt_nasc, *recibos in empresas[index:]:
		result = new_session_senha(cpf, dt_nasc, driver)

		if not isinstance(result, str):
			text = gera_senha_receita(recibos, *result)
			result[1].close()
		else:
			text = result

		print('>>>', f'{cpf};{text}')
		escreve_relatorio_csv(f'{cpf};{text}')

	return True

if __name__ == '__main__':
	run()