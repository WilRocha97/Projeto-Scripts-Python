from bs4 import BeautifulSoup
from Dados import empresas
import sys
sys.path.append("..")
from comum import new_login, new_session, escreve_relatorio_csv, \
atualiza_info, salvar_arquivo, time_execution


def consulta_atrasos(s):
	print('>>> Consultando parcelas em atraso')

	headers = {
        'Accept-Encoding': 'gzip, deflate, sdch',
        'Accept-Language': 'en-US,en;q=0.8',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
    }

	url_base = 'https://cav.receita.fazenda.gov.br/Servicos/ATSPO/parcweb.app'
	url_inicial = f'{url_base}/parcweb_gerenciador.asp'
	url_extrato = f'{url_base}/ExtratoRelacaoProcessos.asp'

	s.get(url_inicial, headers=headers)

	response = s.get(url_extrato, headers=headers)
	soup = BeautifulSoup(response.content, 'html.parser')
	msg = soup.find('form', attrs={'id': 'form1'})
	if not msg:
		if not soup.find('td', attrs={'class':'TDTrataParWeb'}):
			return 'Nao pode ser consultado'

		elem = soup.select("body > div:nth-child(5) > table")[0]
		for tr in elem.find_all('tr')[1:]:
			compl = tr.find_all('td')[-1].a.get('href', '')

			response = s.get(f"{url_base}/{compl}", headers=headers)
			aux_soup = BeautifulSoup(response.content, 'html.parser')
			table = aux_soup.select("body > div:nth-child(5) > table")[0]

			for aux_tr in table.find_all('tr'):
				parcelas_atraso = aux_tr.find_all('td')
				if not parcelas_atraso: continue

				n_parc = parcelas_atraso[4].text
				if int(n_parc) != 0: return "parcelas em atraso"
			else:
				return "parcelamento sem atrasos"

			break

		return 'Sem parcelamento'	
	
	# Isso cobre os casos que não existem parcelamentos para consultar
	return msg.center.text.strip()


@time_execution
def run():
	prev_cert, s = None, None
	total, lista_empresas = len(empresas), empresas

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
			situacao = consulta_atrasos(s)
		else:
			situacao = {res["Value"]}
		
		texto = f'{cnpj};{situacao}'
		escreve_relatorio_csv(texto)
		print(f'>>> {texto}\n\n', end='')

	else:
		if s != None: s.close()

	#header = ""
	#escreve_header_csv(header)


if __name__ == '__main__':
	run()