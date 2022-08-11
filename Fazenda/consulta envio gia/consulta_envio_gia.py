# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from sys import path
import os

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from fazenda_comum import new_session_fazenda
from comum_comum import time_execution, open_lista_dados, escreve_relatorio_csv, where_to_start, escreve_header_csv


def create_pdf(nome, caminho=r'execucao\docs', content=None, soup=None):
    if not soup:
        soup = BeautifulSoup(content, 'html.parser')
    os.makedirs(caminho, exist_ok=True)
    # Remove atributo href das tags, links e urls não serão usados
    for tag in soup.find_all(href=True):
        tag['href'] = ''

    # Remove tags img e script, não serão usados
    _ = list(tag.extract() for tag in soup.find_all('img'))
    _ = list(tag.extract() for tag in soup.find_all('script'))

    with open(os.path.join(caminho, f'{nome}.pdf'), 'w+b') as pdf:
        pisa.showLogging()
        pisa.CreatePDF(str(soup), pdf)


def consulta_gia(ni, mes, ano, session, session_id):
    url_base = "https://cert01.fazenda.sp.gov.br"
    url = f"{url_base}/novaGiaWEB/consultaIe.gia"

    data = {
        'SID': session_id, 'servico': 'GIA', 'method': 'consultaIe',
        'ie': '{}.{}.{}.{}'.format(ni[:3], ni[3:6], ni[6:9], ni[9:]),
        'refInicialMes': mes, 'refInicialAno': ano,
        'refFinalMes': mes, 'refFinalAno': ano,
    }

    jsessionid = session.cookies.get_dict().get('JSESSIONID', '')
    res = session.post(f'{url};jsessionid={jsessionid}', data, verify=False)
    soup = BeautifulSoup(res.content, 'html.parser')

    erro = soup.find('td', attrs={'class': "RESULTADO-ERRO"})
    if erro:
        return ' '.join(erro.text.split())

    valores = soup.find_all('td', attrs={'class': 'RESULTADO-VALOR'})
    valores = list(map(lambda x: x.text.split(':')[1].strip(), valores))
    valores = list(valores[i] for i in (1, 3, 4, 5, 9, 10))

    valores[0] = "".join(i for i in valores[0] if i.isdigit())
    valores[2] = valores[2].split(" - ")[0]

    pdf = f'consulta gia {valores[0]} {mes}{ano}'
    create_pdf(nome=pdf, content=res.content)

    return ';'.join(valores)


@time_execution
def run():
    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp:
        return False
    comp = comp.split('/')

    empresas = open_lista_dados()
    if not empresas:
        return False

    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    for ni, user, pwd, tipo in empresas[index:]:
        res = new_session_fazenda(ni, user, pwd, tipo)
        if not isinstance(res, str):
            text = consulta_gia(ni, *comp, *res)
            res[0].close()
        else:
            text = res

        print('>>>', text)
        escreve_relatorio_csv(f'{ni};{text}')

    escreve_header_csv('inscricao;cnpj;ref;dt. entrega;tipo;chave;protocolo')
    return True


if __name__ == '__main__':
    run()
