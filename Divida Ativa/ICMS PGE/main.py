from requests import Session
from bs4 import BeautifulSoup
from datetime import datetime
from Dados import empresas
import os


def salva_registro(cnpj, texto):
    nome = cnpj+' - Ocorrencias.txt'
    try:
        registro = open(os.path.join('guias', nome), 'a')
    except:
        registro = open(os.path.join('guias', nome), 'w')
    registro.write(texto)
    registro.close()


def verifica_mensagem(html, cnpj, cda):
    soup = BeautifulSoup(html, 'html.parser')
    mensagem = soup.find('table', attrs={'id': 'messages'}).find('td').findAll('span')
    try:
        texto = mensagem.text+f'({cda})\n'
        salva_registro(cnpj, mensagem[1].text)
        return True
    except:
        return False


def verifica_parcela(html, cnpj, cda, data):
    soup = BeautifulSoup(html, 'html.parser')
    # ATUALIZAR OS ID'S DE ACORDO COM A TABELA DE PARCELAS SELECIONADAS
    try:
        tabela = soup.find('tbody', attrs={'id': 'gareForm:j_id379:tb'}).findAll('tr')
    except:
        print('Falha ao procurar as Parcelas')
        return False

    for index, linha in enumerate(tabela):
        vencimento = linha.find('td', attrs={'id': f'gareForm:j_id379:{index}:j_id395'}).text
        if vencimento != data:
            parcela = linha.find('td', attrs={'id': f'gareForm:j_id379:{index}:j_id393'}).text
            situacao = linha.find('td', attrs={'id': f'gareForm:j_id379:{index}:j_id397'}).text
            texto = f'CDA: {cda} > '
            texto += f'Parcela {parcela} - Vencimento: {vencimento} - '
            texto += f'Situação: {situacao}\n'
            salva_registro(cnpj, texto)

    return True if tabela else False


def download_guia(cnpj, cda, data):
    url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/pagamento/gareParcelamento.jsf'

    with Session() as s:
        pagina = s.get(url)
        soup = BeautifulSoup(pagina.text, 'html.parser')
        ViewState = soup.find('input', attrs={'id':'javax.faces.ViewState'})['value']
        
        info = {
            'AJAXREQUEST': '_viewRoot', 'gareForm': 'gareForm',
            'gareForm:inputCDA': cda, 'gareForm:modalSelecionarDebitoOpenedState': '',
            'javax.faces.ViewState': ViewState,
            'gareForm:buttonAvancar': 'gareForm:buttonAvancar'
        }

        pagina = s.post(url, info)
        if verifica_mensagem(pagina.text, cnpj, cda):
            print(f'Não existe parcelamento para a CDA informada. ({cda})')
            return False

        info.pop('gareForm:inputCDA')
        pagina = s.post(url, info)
        if verifica_mensagem(pagina.text, cnpj, cda):
            print(f'As parcelas não podem ser emitidas pois excedeu o limite de 90 dias. ({cda})')
            return False

        info = {
            'AJAXREQUEST': '_viewRoot', 'gareForm': 'gareForm',
            'gareForm:j_id251:22:j_id266': 'on', 'gareForm:j_id251:23:j_id266': 'on',
            'gareForm:j_id300:j_id305': '',
            'gareForm:modalSelecionarDebitoOpenedState': '',
            'javax.faces.ViewState': ViewState,
            'gareForm:buttonAvancar': 'gareForm:buttonAvancar'
        }
        pagina = s.post(url, info)

        if not verifica_parcela(pagina.text, cnpj, cda, data):
            texto = f'Não foram encontradas parcelas para a competência solicitada ({cda})\n'
            salva_registro(cnpj, texto)
            return False

        info.pop('gareForm:j_id251:22:j_id266')
        info.pop('gareForm:j_id251:23:j_id266')
        info['gareForm:j_id379:0:j_id392'] = 'on'
        #info['gareForm:j_id375:1:j_id384'] = 'on'
        s.post(url, info)

        info = {
            'gareForm': 'gareForm', 'gareForm:j_id464': 'Download',
            'gareForm:modalSelecionarDebitoOpenedState': '',
            'javax.faces.ViewState': ViewState
        }
        salvar = s.post(url, info)
        filename = salvar.headers.get('content-disposition', '')

        if filename:
            n = data.split('/')
            str_venc = '-'.join(n)
            nome =  ";".join([
                cnpj, 'IMP_FISCAL',  str_venc,
                'Parcelamento ICMS', n[1], 
                n[2], 'Parcelamento', 'ICMS PGE 1 - ' + cda
            ]) + ".pdf"
            arquivo = open(os.path.join('guias', nome), 'wb')
            for parte in salvar.iter_content(100000):
                arquivo.write(parte)
            arquivo.close()
            print(f"Arquivo {nome} salvo")
        else:
            print('Não foi encontrado dados para salvar')


if __name__ == '__main__':
    inicio = datetime.now()
    os.makedirs('guias', exist_ok=True)
    data = '28/02/2022'

    lista_auxiliar = [
        ('13296548000172', '1265533311'),
    ]

    lista = empresas
    #lista = lista_auxiliar
    for empresa in lista:
        print(f'>>> Iniciando CDA {empresa[1]}')
        download_guia(empresa[0], empresa[1], data)
    print('>>> Programa finalizado. Tempo total: ',(datetime.now()-inicio))
