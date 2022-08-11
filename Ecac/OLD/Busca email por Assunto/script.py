# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from pyautogui import prompt
from Dados import empresas
import sys, os
sys.path.append('..')
from comum import new_session, new_login, escreve_relatorio_csv, \
time_execution


def procurar_mensagen(cnpj, s, assunto):
    print('>>> Procurando mensagens')

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
    url = 'https://cav.receita.fazenda.gov.br/Servicos/ATSDR/CaixaPostal.app/Action/'
    url_css = 'https://cav.receita.fazenda.gov.br/Servicos/ATSDR/CaixaPostal.app/Content/cxpostal.css'

    s.get(url_home, headers=headers)

    pagina = s.get(url+'ListarMensagemAction.aspx', headers=headers, verify=False)
    if pagina.status_code != 200:
        print('>>> Não carregou')
        escreve_relatorio_csv(f'{cnpj};Erro ao carregar página')
        return False

    soup = BeautifulSoup(pagina.content, 'html.parser')
    try:
        tabela_msg = soup.find('table', attrs={'id':'gridMensagens'}).findAll('tr')
    except:
        print('>>> Não tem mensagens')
        escreve_relatorio_csv(f'{cnpj};Sem mensagens na caixa de entrada')
        return False

    for index, linha in enumerate(tabela_msg, 1):
        colunas = linha.findAll('td')
        if assunto in colunas[4].text.strip().lower():
            css = s.get(url_css, headers=headers)
            css = css.text

            id_parc = "gridMensagens_ctl" + str(index).rjust(2, '0')
            objeto_m = soup.find('a', attrs={'id': id_parc + '_nmAssunto'}).get('href')
            pagina = s.get(url + objeto_m, headers=headers)
            soup = BeautifulSoup(pagina.content, 'html.parser')
            
            link_css = soup.new_tag('style', type="text/css")
            link_css.string = css
            soup.find('link').append(link_css)

            nome_arq = f'{cnpj} - mensagem.html'
            with open(os.path.join('documentos', nome_arq), 'w', encoding="latin-1", errors='ignore') as arquivo:
                arquivo.write(soup.prettify())
            print('>>> Mensagem salva')
            escreve_relatorio_csv(f'{cnpj};Possui notificacao')
            break
    else:
        print('>>> Não tem mensagens')
        escreve_relatorio_csv(f'{cnpj};Sem notificacao')

    return True


@time_execution
def run():
    prev_cert, s = [None] * 2 
    os.makedirs('documentos', exist_ok=True)

    lista_empresas = empresas
    total = len(lista_empresas)
    assunto = prompt(title="Busca por assunto", text="Digite o assunto exatamente como esta no ecac:")
    if not assunto:
        print("Execução cancelada")
        return False

    print("Pesquisando por", assunto)
    for index, empresa in enumerate(lista_empresas, 1):
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
            procurar_mensagen(cnpj, s, assunto)
        else:
            print('>>> Não logou')
            escreve_relatorio_csv(f'{cnpj};{res["Value"]}')

        print(f'>>> Busca finalizada - {index}/{total}\n')
    else:
        if s != None: s.close()


if __name__ == '__main__':
    run()

