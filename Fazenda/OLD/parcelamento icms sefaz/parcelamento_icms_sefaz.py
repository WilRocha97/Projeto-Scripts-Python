# -*- coding: utf-8 -*-
from datetime import datetime
from bs4 import BeautifulSoup
from pyautogui import prompt
from Dados import empresas
import sys
sys.path.append("..")
from comum import login_usuario, atualiza_info, salvar_arquivo, \
escreve_relatorio_csv, escreve_header_csv


def lista_acordos(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    tipo_parc = [
        'Parcelamento com Pagamento à Menor',
        'Parcelamento em Andamento'
    ]

    try:
        attrs = {'id':'MainContent_gvListaPedidoParcelado'}
        linhas = soup.find('table', attrs=attrs)
        linhas = linhas.find('tbody').findAll('tr')
    except:
        return []

    acordos = []
    for index, linha in enumerate(linhas):
        colunas = linha.findAll('td')
        situacao = colunas[3].text.strip()
        if situacao not in tipo_parc: continue

        acordo = colunas[0].text.strip()
        n_acordo = 'ctl' + str(index + 2).rjust(2, '0')
        acordos.append((n_acordo, acordo))

    return acordos


def verifica_mensagem(pagina):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    try:
        attrs = {'id': 'MainContent_lblMensagemAlerta'}
        texto = soup.find('span', attrs=attrs).text.strip()
    except Exception as e:
        texto = 'Excecao ao recuperar mensagem'

    return texto if texto and texto != 'Label' else ''


def verifica_guias(pagina, comp):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    comp = datetime.strptime(comp, '%d/%m/%Y')

    try:
        attrs = {'id':'MainContent_gwListaParcelas'}
        linhas = soup.find('table', attrs=attrs)
        linhas = linhas.findAll('tr')
    except:
        return 'Erro ao carregar parcelas'

    link, atraso = '', []
    for linha in linhas[2:-1]:
        colunas = linha.findAll('td')
        imprimir = colunas[0].find('a')
        
        if not imprimir: continue

        parcela = colunas[1].text.strip()
        venc = colunas[2].text.strip()
        venc = datetime.strptime(venc, '%d/%m/%Y')

        if venc == comp:
            link = imprimir.get('href', '').rstrip("','')")
            link = 'ct' + link.lstrip("javascript:__doPostBack('")
        elif venc < comp:
            atraso.append(venc.strftime('%d/%m/%Y'))
        else:
            break

    return (link, atraso)


def consulta_parc(cnpj, session, s_id, venc):
    url_base = 'https://www10.fazenda.sp.gov.br/ParcelamentoIcms/Pages'
    url_consulta = f'{url_base}/ConsultarAlterarParcelamento.aspx'
    url_impressao = f'{url_base}/Impressao.aspx'

    # o login quando feito com usuario e senha do contribuinte
    # permite a busca apenas utilizando insc. estadual
    if len(cnpj) == 12:
        ident = 'IE'
        f_cnpj = f'{cnpj[:3]}.{cnpj[3:6]}.{cnpj[6:9]}.{cnpj[9:]}'
    else:
        ident = 'CNPJ'
        f_cnpj = f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
    
    id_base = 'ctl00$MainContent$'
    pagina = session.get(f'{url_consulta}?SID={s_id}')
    state, generator, validation = atualiza_info(pagina)
    info = {
        '__EVENTTARGET': f'{id_base}btnConsultar', '__EVENTARGUMENT': '',
        '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
        '__EVENTVALIDATION': validation, f'{id_base}ddlContribuinte': ident,
        f'{id_base}txtCnjpIE': f_cnpj
    }

    pagina = session.post(f'{url_consulta}?SID={s_id}', info)
    msg = verifica_mensagem(pagina)
    if msg: return f'{cnpj};{msg}'

    acordos = lista_acordos(pagina)
    if not acordos:        
        return f'{cnpj};Não foram encontrados acordos ativos'

    textos = []
    for index, acordo in acordos:
        print(f'>>> Verificando acordo {acordo}')
        sit_guia = 'Guia atual não encontrada'
        sit_atraso = 'Sem guias em atraso'

        state, generator, validation = atualiza_info(pagina)
        target = f'{id_base}gvListaPedidoParcelado${index}$lnkNumeroParcelamento'
        info = {
            '__EVENTTARGET': target, '__EVENTARGUMENT': '',
            '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
            '__EVENTVALIDATION': validation
        }

        parcelas = session.post(f'{url_consulta}?SID={s_id}', info)
        state, generator, validation = atualiza_info(parcelas)
        target, atraso = verifica_guias(parcelas, venc)
        if atraso: sit_atraso = ', '.join(atraso)

        if not target:
            textos.append(f'{cnpj};{acordo};{sit_guia};{sit_atraso}')
            continue

        info = {
            '__EVENTTARGET': target, '__EVENTARGUMENT': '',
            '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': generator,
            '__EVENTVALIDATION': validation
        }

        session.post(f'{url_consulta}?SID={s_id}', info)
        pagina_externa = session.get(url_impressao)

        html = pagina_externa.text
        i, f= html.find('ExportUrlBase') + 16, html.find('FixedTableId') - 3
        compl_url = html[i:f].replace('\\u0026', '&')

        url_download = f'https://www10.fazenda.sp.gov.br{compl_url}PDF'
        salvar = session.get(url_download)
        filename = salvar.headers.get('content-disposition', '')
        if filename:
            str_venc = venc.replace('/', '-')
            list_venc = venc.split('/')

            nome = (
                f'{cnpj};IMP_FISCAL;{str_venc};Parcelamento ICMS;' +
                f'{list_venc[1]};{list_venc[2]};Parcelamento;' +
                f'ICMS Sefaz1 - {acordo}.pdf'
            )
            salvar_arquivo(nome, salvar)
            sit_guia = 'Guia atual disponivel'

        textos.append(f'{cnpj};{acordo};{sit_guia};{sit_atraso}')

    return '\n'.join(textos)


def run():
    comeco = datetime.now()
    print("Execução iniciada as: ", comeco)
    print("Insira a data de vencimento na caixa de texto... ", end='')

    total, lista_empresas = len(empresas), empresas
    venc = prompt("Digite o vencimento das guias 00/00/0000:")
    if not venc:
        print("Execução abortada")
        return False

    print(venc)
    for i, empresa in enumerate(lista_empresas):
        cnpj = empresa[0]

        print(i, total)
        s, s_id = login_usuario(*empresa)

        if s:
            texto = consulta_parc(cnpj, s, s_id, venc)
            s.close()
        else:
            texto = f'{cnpj};{s_id}'

        escreve_relatorio_csv(texto)
        print(f'>>> {texto}\n\n', end='')

    escreve_header_csv('cnpj;acordo;sit. guia do mes atual;sit. guias em atraso')
    print("Tempo de execução: ", datetime.now() - comeco)
    return True


if __name__ == '__main__':
    run()