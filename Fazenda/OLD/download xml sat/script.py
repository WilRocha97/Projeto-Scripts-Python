# -*- coding: utf-8 -*-
from datetime import datetime, timedelta
from calendar import monthrange
from selenium import webdriver
from bs4 import BeautifulSoup
from requests import Session
from time import sleep
import sys, os, re
sys.path.append("..")
from comum import pfx_to_pem


'''
 Automatiza o download de lotes de sat em .xml no site da sefaz

 *Antes de iniciar o processo é necessário alterar o intervalo da busca na variavel
 _datas no formato ("di/mi/yyyi", "df/mf/yyyf") sendo a data incial e a data final
 respectivamente

 **A planilha de dados desse script segue o formato:
 empresas = [
	("cnpj", "caminho_certificado", "senha_certificado"),
	...
 ]

'''

_datas = ("01/01/2020", "30/12/2020")


def new_session(cert, senha):
	print(">>> Logando com", cert)
	url_login = "https://satsp.fazenda.sp.gov.br/COMSAT/Account/LoginSSL.aspx"

	with pfx_to_pem(cert, senha) as cert:
		file = os.path.join('..', 'phantomjs.exe')
		args = [f'--ssl-client-certificate-file={cert}']
		driver = webdriver.PhantomJS(file, service_args=args)
		driver.set_window_size(1000, 900)
		driver.delete_all_cookies()
	
	for i in range(5):
		try:
			driver.get(url_login)
			sleep(1)
			driver.find_element_by_id("conteudo_rbtContabilista").click()
			sleep(1)
			driver.find_element_by_id("conteudo_imgCertificado").click()
			break
		except Exception as e:
			print("\tnova tentativa")
	else:
		driver.quit()
		return False

	cookies = driver.get_cookies()
	driver.quit()

	session = Session()
	for cookie in cookies:
		session.cookies.set(cookie['name'], cookie['value'])

	return session


def att_post_data(content):
	soup = BeautifulSoup(content, "html.parser")
	state, validation, viewgenerator = [None] * 3

	try:
		state = soup.find("input", attrs={"id": "__VIEWSTATE"}).get("value")
		validation = soup.find("input", attrs={"id": "__EVENTVALIDATION"}).get("value")
		viewgenerator = soup.find("input", attrs={"id": "__VIEWSTATEGENERATOR"}).get("value")
	except:
		raise Exception(state, validation, viewgenerator)

	return state, viewgenerator, validation


def divide_meses(datas):
	dt_i = datetime.strptime(datas[0], '%d/%m/%Y')
	dt_f = datetime.strptime(datas[1], '%d/%m/%Y')

	lista = []
	while True:
		year, month = dt_i.year, dt_i.month
		last_day = monthrange(year, month)[1]
		tmp_f = datetime(year, month, last_day)

		str_i = datetime.strftime(dt_i, '%d/%m/%Y')
		if tmp_f >= dt_f:
			str_f = datetime.strftime(dt_f, '%d/%m/%Y')
			lista.append((str_i, str_f))
			break
		
		str_f = datetime.strftime(tmp_f, '%d/%m/%Y')
		lista.append((str_i, str_f))
		dt_i = tmp_f + timedelta(days=1)

	return lista


def download_lotes(response):
	filename = response.headers.get('content-disposition', '')

	if filename:
		nome = filename.split('=')[-1]
		arquivo = open(os.path.join('xml', nome), 'wb')
		for parte in response.iter_content(100000):
			arquivo.write(parte)
		arquivo.close()
		print(">>> Baixou", nome)

	return True


def pesquisa_lotes(session, esat, datas):
	url_base = "https://satsp.fazenda.sp.gov.br/COMSAT/Private"
	url_fsat = f"{url_base}/ConsultaCfeSemErros/ConsultarCfeSemErro.aspx"

	response = session.get(url_fsat)
	viewstate, generator, validation = att_post_data(response.content)

	lotes_form = {
		"ToolkitScriptManager1_HiddenField": ";;AjaxControlToolkit, Version=4.1.50508.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:pt-BR:0c8c847b-b611-49a7-8e75-2196aa6e72fa:de1feab2:fcf0e993:f2c8e708:720a52bf:f9cec9bc:589eaa30:698129cf:fb9b4c57:ccb96cf9",
		"__EVENTTARGET": "", "__EVENTARGUMENT": "", "__VIEWSTATE": viewstate,
		"__VIEWSTATEGENERATOR": generator, "__SCROLLPOSITIONX": "0",
		"__SCROLLPOSITIONY": "0", "__EVENTVALIDATION": validation,
		"ctl00$conteudo$txtNumeroSerie": esat, "ctl00$conteudo$txtChaveAcesso": "",
		"ctl00$conteudo$txtNumeroCupom": "", "ctl00$conteudo$ddlTipoCupom": "0",
		"ctl00$conteudo$ddlSituacao": "0", "ctl00$conteudo$txtDataInicio": datas[0],
		"ctl00$conteudo$txtHoraInicio": "00:00:00", "ctl00$conteudo$txtDataTermino": datas[1],
		"ctl00$conteudo$txtHoraTermino": "23:59:59", "ctl00$conteudo$btnPesquisar": "Pesquisar",
	}
	
	prev, new_search = None, True
	while new_search:
		print("-" * 120)
		print(">>>Buscando")
		response = session.post(url_fsat, data=lotes_form)
		viewstate, generator, validation = att_post_data(response.content)
		if 'Para visualizar os demais registros refine a busca.' not in response.text:
			new_search = False

		page_form = lotes_form.copy()
		page_form.pop("ctl00$conteudo$btnPesquisar")

		soup = BeautifulSoup(response.content, 'html.parser')
		list_page = soup.find('select', attrs={'id': 'conteudo_ddlPages'})
		if not list_page:
			print("Nenhum Resultado")
			break
		
		total_pages = len(list_page.find_all('option'))
		for i in range(1, total_pages + 1):
			page_form["__VIEWSTATE"] = viewstate
			page_form["__EVENTVALIDATION"] = validation
			page_form["__VIEWSTATEGENERATOR"] = generator
			page_form["ctl00$conteudo$ddlPages"] = str(i)

			table = soup.find('table', attrs={'id':'conteudo_grvConsultaCfeSemErros'})
			for row in table.find_all('tr')[1:]:
				lote = row.find_all('a')[-1]
				n_lote = lote.get_text().strip()
				if n_lote == prev: continue

				prev = n_lote
				lote = lote.get("href").lstrip("javascript:__doPostBack('").rstrip("','')")

				page_form["__EVENTTARGET"] = 'ct' + lote
				download = session.post(url_fsat, data=page_form)
				download_lotes(download)

			if i == total_pages:
				id_tag = 'conteudo_grvConsultaCfeSemErros_lblDataProcessamento_9'
				tag = soup.find('span', attrs={'id':id_tag})
				if tag: new_dt, new_hr = tag.get_text().strip().split()
				continue
			
			# POST - trocar de pagina
			page_form['__EVENTTARGET'] = 'ctl00$conteudo$lnkBtnProxima'
			response = session.post(url_fsat, data=page_form)
			soup = BeautifulSoup(response.content, 'html.parser')
			viewstate, generator, validation = att_post_data(response.content)
			# POST - trocar de pagina

		if new_search:
			lotes_form['__VIEWSTATE'] = viewstate
			lotes_form['__EVENTVALIDATION'] = validation
			lotes_form['__VIEWSTATEGENERATOR'] = generator
			lotes_form['ctl00$conteudo$txtDataInicio'] = new_dt
			lotes_form['ctl00$conteudo$txtHoraInicio'] = new_hr



def consulta_sat(cnpj, session, datas):
	url_base = "https://satsp.fazenda.sp.gov.br/COMSAT/Private"
	url_esat = f"{url_base}/VisualizarEquipamentoSat/VisualizarEquipamentoSAT.aspx"

	#Consultar numero dos equipamentos sat ativos
	response = session.get(url_esat)
	viewstate, generator, validation = att_post_data(response.content)

	print("'")
	data = {
		"ToolkitScriptManager1_HiddenField": "", "__EVENTTARGET": "",
		"__EVENTARGUMENT": "", "__LASTFOCUS": "", "__VIEWSTATE": viewstate,
		"__VIEWSTATEGENERATOR": generator, "__SCROLLPOSITIONX": "0",
		"__SCROLLPOSITIONY": "0", "__EVENTVALIDATION": validation,
		"ctl00$conteudo$txtNumeroSerie": "",
		"ctl00$conteudo$ddlFabricante": "-9223372036854775808",
		"ctl00$conteudo$txtCnpjSofwareHouse": "",
		"ctl00$conteudo$ddlSituacao": "10",
		"ctl00$conteudo$btnPesquisar": "Pesquisar",
	}

	for i in range(5):
		response = session.post(url_esat, data=data)
		soup = BeautifulSoup(response.content, "html.parser")
		tabela = soup.find("table", attrs={"id": "conteudo_grvPesquisaEquipamento"})
		if tabela: break

		viewstate, generator, validation = att_post_data(response.content)
		data["__VIEWSTATE"] = viewstate
		data["__EVENTVALIDATION"] = validation
		data["__VIEWSTATEGENERATOR"] = generator
	else:
		raise Exception("Empresa não tem Equipamento SAT ativo")

	n_esats = []
	for n, _ in enumerate(tabela.find_all("tr")):
		a_id = f"conteudo_grvPesquisaEquipamento_lblNumeroSerie_{n}"
		link = tabela.find("a", attrs={"id": a_id})
		if link: n_esats.append(link.get_text())

	#Para cada numero de equipamento sat, pesquisar pelos lotes
	for esat in n_esats:
		for data in divide_meses(datas):
			print("\n>>>Consultando empresa {} equipamento {} na data {}\n".format(cnpj, esat, data))
			pesquisa_lotes(session, esat, data)

	return True


def new_login(cnpj, session):
	print(">>> Acessando", cnpj)
	url_base = "https://satsp.fazenda.sp.gov.br/COMSAT/Private"
	url_cnpj = f"{url_base}/SelecionarCNPJ/SelecionarCNPJContribuinte.aspx"

	response = session.get(url_cnpj, verify=False)
	f_cnpj = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
	viewstate, generator, validation = att_post_data(response.content)

	data = {
		"ToolkitScriptManager1_HiddenField": "", "__EVENTTARGET": "",
		"__EVENTARGUMENT": "", "__LASTFOCUS": "", "__VIEWSTATE": viewstate,
		"__VIEWSTATEGENERATOR": generator, "__SCROLLPOSITIONX": "0",
		"__SCROLLPOSITIONY": "0", "__EVENTVALIDATION": validation,
		"ctl00$conteudo$txtCNPJ_ContribuinteNro": f_cnpj[:10],
		"ctl00$conteudo$txtCNPJ_ContribuinteFilial": f_cnpj[11:],
		"ctl00$conteudo$btnPesquisar": "Pesquisar",
		"ctl00$conteudo$ddlPages": "1"
	}

	response = session.post(url_cnpj, data=data)
	viewstate, generator, validation = att_post_data(response.content)

	updated_data = {
		"__EVENTTARGET": "ctl00$conteudo$gridCNPJ$ctl02$lnkCNPJ",
		"__VIEWSTATE": viewstate, "__VIEWSTATEGENERATOR": generator,
		"__EVENTVALIDATION": validation, "__ASYNCPOST": "false",
		"ctl00$conteudo$gridCNPJ$ctl02$lnkCNPJ": "lnkCNPJ",
		"ctl00$conteudo$cnpjCpfDLL": "CNPJ",
	}
	data.update(updated_data)
	data.pop('ctl00$conteudo$btnPesquisar')
	data.pop('ctl00$conteudo$ddlPages')
	session.post(url_cnpj, data=data)


def run():
	comeco = datetime.now()
    print("Execução iniciada as: ", comeco)

    # Testes
	empresas = [
		("69092005000198", r"..\certificados\CERT RPEM 35586086.pfx", "35586086"),
	]
	# Testes

	prev_cert, session = [None] * 2
	lista_empresas = empresas
	os.makedirs('xml', exist_ok=True)

	for cnpj, *cert in lista_empresas:

		if cert != prev_cert:
			if session: session.close()
			session = new_session(*cert)
			prev_cert = cert
			if not session:
				prev_cert = None
				continue

		try:
			new_login(cnpj, session)
		except:
			print(">>>Não logou")
			continue

		consulta_sat(cnpj, session, _datas)
	else:
        if s != None: s.close()
	
	print("Tempo de execução: ", datetime.now() - comeco)


if __name__ == '__main__':
	run()