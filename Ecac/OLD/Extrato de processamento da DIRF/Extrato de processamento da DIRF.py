# -*- coding: utf-8 -*-
import sys, os, re
from bs4 import BeautifulSoup
from datetime import datetime
from Dados import empresas
from urllib import parse
sys.path.append('..')
from comum import new_session, new_login, escreve_relatorio_csv


def extrato():
    # dicionários
    keys3 = {'ni': cnpj, 'tipo_ni': 2, '_': ''}

    # urls
    urlg1 = 'https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx?id=12&origem=menu'
    urlg2 = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/default.asp'
    urlg3 = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/DirfP002/DirfP002.asp'
    urlp1 = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/DirfC100/DirfC100.asp?Width=1366&Height=768'
    urlp3 = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/dirfc500/dirfc500.asp'

    # entrar no site dos extratos da DIRF
    s.get(urlg1, verify=False)
    s.get(urlg2)
    s.get(urlg3)
    pagina = s.post(urlp3, data=keys3)
    soup = BeautifulSoup(pagina.content, 'html.parser')

    # Pegar o nome da empresa para usar nos próximos requests
    padraozinho = re.compile(r'CNPJ.+\n.+\n.+\n.+\n.+\n\s+(.+)')
    nome = s.post(urlp1)
    if pagina.status_code != 200:
        print('*** ERRO NA APLICAÇÃO ***\n')
        escreve_relatorio_csv(';'.join([certif, cnpj, '...', '...', '...', 'Ocorreu um erro interno na aplicação.']))
        return False
    soup_nome = BeautifulSoup(nome.content, 'html.parser')
    nome = soup_nome.prettify()
    matchzinho = padraozinho.search(nome)

    nome = matchzinho.group(1)
    print(nome)

    # Verificar se existe extrato
    soup_comp = BeautifulSoup(pagina.content, 'html.parser')
    soup_comp = str(soup_comp)
    resultado = re.compile(r'Programa Dirf x Darf Exercício')
    resultado = resultado.search(soup_comp)
    if not resultado:
        print('** Não possui extrato de processamento da DIRF **\n')
        escreve_relatorio_csv(';'.join([certif, cnpj, nome, '...', '...', 'Não possui extrato de processamento da DIRF']))
        return False

    # Pegar quantas competências da DIRF existem para a empresa
    anos = soup.get_text('\n')
    anos = anos.replace('Programa Dirf x Darf Exercício : ', '')
    anos = anos.split()

    for ano in anos:
        ano = int(ano)
        ano = ano - 1
        print(ano + 1)

        # Consultar
        urlp = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/DirfC530/DirfC530.asp?'
        keys = {'ni': cnpj,
                'tipo_ni': 2,
                'no_ni': nome,
                'ano_calendario': ano}
        pagina = s.post(urlp, data=keys)
        if pagina.status_code != 200:
            print('*** ERRO NA APLICAÇÃO ***\n')
            anotexto = ano
            escreve_relatorio_csv(';'.join([certif, cnpj, nome, str(anotexto), '...', 'Ocorreu um erro interno na aplicação.']))
            continue

        # Pegar quantas guias a empresa teve no ano
        soup = BeautifulSoup(pagina.content, 'html.parser')
        padraozinho = re.compile(r'darf\.asp\?(.+)\"')
        soup = str(soup)
        links = padraozinho.findall(soup)

        for link in links:

            # Montar a url para entrar nos pagmentos da guia
            url_base = 'https://cav.receita.fazenda.gov.br/Servicos/ATCTA/dirfweb/DirfC570/DirfC570_darf.asp'

            padrao = re.compile(r'cod_receita=(\d+)')
            matchzinho = padrao.search(link)
            cod = matchzinho.group(1)

            padrao = re.compile(r'valor_dirf=(\d+)')
            matchzinho = padrao.search(link)
            valor = matchzinho.group(1)

            padrao = re.compile(r'da_dirf=(\d+)')
            matchzinho = padrao.search(link)
            data = matchzinho.group(1)

            query = {'origem': 'C530',
                     'ano': ano,
                     'ni': cnpj,
                     'tipo_ni': '2',
                     'cod_inf': '208',
                     'cod_receita': cod,
                     'valor_dirf': valor,
                     'drf': '0810400',
                     'da_dirf': data,
                     'nu_intimacao': '',
                     'in_intimado_reintimado': '',
                     'data_entrega': '',
                     'txtNR': '',
                     'in_normal_especial': '1',
                     'no_ni': nome}

            parsed_url = parse.urlparse(url_base)._replace(query=parse.urlencode(query))
            url_login = parse.urlunparse(parsed_url)
            url_login = url_login.replace('+', '%20')

            # Entar nos pagamentos da guia
            pagina = s.get(url_login)
            # guardar o código da página para salvar em um arquivo html
            soup = BeautifulSoup(pagina.content, 'html.parser')
            soup2 = BeautifulSoup(pagina.content, 'html.parser')
            soup2 = str(soup2)

            # Verificar se existe recolhimento
            resultado = re.compile(r'Não foram encontrados recolhimentos')
            resultado = resultado.search(soup2)
            ano = ano + 1

            if not resultado:
                valpadrao = re.compile(r'(\(Pagamentos até\s.+)<\/b><\/b>')
                matchzinho = valpadrao.search(soup2)
                datapag = matchzinho.group(1)
                print(cod + ' - Divergência entre os valores declarados (Dirf) e os recolhidos (Darf).' + datapag + '\n')
                escreve_relatorio_csv(';'.join([certif, cnpj, nome, str(ano), str(cod), 'Divergência entre os valores declarados (Dirf) e os recolhidos (Darf).' + datapag]))

                nome = nome.replace('/', '')
                nome_arq = ' - '.join([nome, cnpj, str(ano), str(cod), 'Divergência entre os valores declarados (Dirf) e os recolhidos (Darf).html'])
                nome_arq = os.path.join('documentos', nome_arq)
                with open(nome_arq, 'w', encoding="latin-1", errors='ignore') as arq:
                    arq.write(soup.prettify())

            else:
                print(cod + ' - Não foram encontrados recolhimentos para este código\n')
                escreve_relatorio_csv(';'.join([certif, cnpj, nome, str(ano), str(cod), 'Não foram encontrados recolhimentos em Darf para este código']))


if __name__ == '__main__':
    comeco = datetime.now()
    print("Execução iniciada as: ", comeco)

    prev_cert, s = [None] * 2
    lista_empresas = empresas
    os.makedirs('documentos', exist_ok=True)

    for count, empresa in enumerate(lista_empresas, start=1):
        cnpjdados, certificado, senha = empresa
        certif = certificado
        certif = certif.split()
        certif = certif[1]
        try:
            # Cria um indice par saber qual linha dos dados está
            indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
            print(indice)

            cnpj, *cert = empresa

            if cert != prev_cert:
                if s: s.close()
                s = new_session(*cert)
                prev_cert = cert
                if not s:
                    prev_cert = None
                    continue

            res = new_login(cnpj, s)
            if res['Key']:
                if not extrato():
                    continue

            else:
                print('** Não logou **')
                escreve_relatorio_csv(f'{certif};{cnpj};"...";"...";"...";{res["Value"]}')

            print('\n', end='')
        except ConnectionError:
            empresas.apend = empresa

            print('ERRO')
            escreve_relatorio_csv(';'.join([certif, cnpjdados, '...', '...', '...', 'ERRO']))
            continue

    else:
        if s != None: s.close()

    print("Tempo de execução: ", datetime.now() - comeco)
