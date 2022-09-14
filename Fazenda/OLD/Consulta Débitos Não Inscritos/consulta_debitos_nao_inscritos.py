# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from requests import Session
from sys import path

path.append(r'..\..\_comum')
from fazenda_comum import _login_ecnd, _atualiza_info
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _download_file, _open_lista_dados, _where_to_start, _indice


def consulta_ipva(cnpj, certificado, senha):
    url = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages/Restrita/PesquisarContribuinte.aspx'
    url_dados = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages/Restrita/VerificarImpedimentosPorCNPJCompleto.aspx'
    url_exibir = f'{url_dados}/Exibir'

    cnpj_format = f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
    cookies = _login_ecnd(certificado, senha)
    with Session() as s:
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])

        for i in range(5):
            try:
                pagina = s.get(url, verify=False)
                state, generator, validation = _atualiza_info(pagina)
                soup = BeautifulSoup(pagina.content, 'html.parser')
                state2 = soup.find('input', attrs={'id': '__VIEWSTATE__'}).get('value')
                break
            except Exception as e:
                if i == 4:
                    print('❌ Erro durante login', e.__class__)
        else:
            print('❌ Falha ao logar na empresa.')
            _escreve_relatorio_csv(f'{cnpj};Falha ao logar na empresa')
            return False

        info = {
            '__EVENTTARGET': 'ctl00$MainContent$btnPesquisar',
            '__EVENTARGUMENT': '', '__VIEWSTATE__': state2, '__VIEWSTATE': state,
            '__EVENTVALIDATION': validation, 'ctl00$MainContent$txtDocumento': cnpj_format,
            'ctl00$MainContent$grupoDocumento': 'radiocnpj'
        }

        s.post(url, info, verify=False)

        pagina = s.get(url_dados, verify=False)
        dados = s.post(url_exibir, json={'msg': cnpj_format}, verify=False)
        dados = dados.json()['d']
        download = False

        if dados['MensagemDeErro']:
            print('❌ Erro na consulta')
            _escreve_relatorio_csv(f'{cnpj};Empresa sem pendências')
            return False

        for item in dados['ListaSumarizacao']:
            if item['Informacao'] not in ['Não há Débitos', 'Não há Pendências']:
                download = True
                break

        if download:
            state, generator, validation = _atualiza_info(pagina)
            soup = BeautifulSoup(pagina.content, 'html.parser')
            state2 = soup.find('input', attrs={'id': '__VIEWSTATE__'}).get('value')
            info = {
                '__EVENTTARGET': 'ctl00$MainContent$lnkImprimirCertidaoBotao2',
                '__EVENTARGUMENT': '',
                '__VIEWSTATE__': state2,
                '__VIEWSTATE': '',
                '__EVENTVALIDATION': validation
            }
            salvar = s.post(url_dados, info, verify=False)
            nome = f'{cnpj} - Cert Negativa de Debts Nao Inscritos.pdf'
            _download_file(nome, salvar)
            print('❗ Com pendências.')
            texto = f'{cnpj};Com pendências'
        else:
            print('✔ Empresa sem pendências.')
            texto = f'{cnpj};Empresa sem pendências'

        _escreve_relatorio_csv(texto)


@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)

        consulta_ipva(*empresa)


if __name__ == '__main__':
    run()
