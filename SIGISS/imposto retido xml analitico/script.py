# -*- coding: utf-8 -*-
# ! python3
from dateutil.relativedelta import *
from datetime import datetime, date
from requests import Session
from pyautogui import prompt
from Dados import empresas
import xmltodict, os


def escreve_relatorio_csv(texto):
    try:
        arquivo = open("Resumo.csv", 'a')
    except:
        arquivo = open("Resumo.csv", 'w')
    arquivo.write(texto+'\n')
    arquivo.close()


def analisa_xml(arq):
    nome = arq[:-4]
    with open(os.path.join('documentos', arq), 'r', encoding='utf-8') as fd:
        doc = xmltodict.parse(fd.read())

    nota = doc["NFES"]["NOTA_FISCAL"]
    if type(nota) is list: 
        resumo = {}
        for n in nota:
            if n["cancelada"] == 'S': continue
            dados = []
            dados.append(n["PRESTADOR"]["cnpj_cpf"])
            dados.append(n["PRESTADOR"]["razao_social"])
            dados.append(n["DESTINATARIO"]["cnpj_cpf"])
            dados.append(n["DESTINATARIO"]["razao_social"])
            try:
                inss = float(resumo[dados[2]][0].replace(',','.'))
            except:
                inss = 0.00
            inss_xml = n["impostos"]["valor_inss"] or '0,00'
            inss += float(inss_xml.replace('.', '').replace(',', '.').replace('R$ ', ''))
            dados.append(str(inss).replace('.', ','))
            resumo[dados[2]] = str(inss), dados

        for r in resumo.values():
            escreve_relatorio_csv(';'.join(str(d) for d in r[1]))

    else:
        if nota["cancelada"] == 'S': return True
        dados = []
        dados.append(nota["PRESTADOR"]["cnpj_cpf"])
        dados.append(nota["PRESTADOR"]["razao_social"])
        dados.append(nota["DESTINATARIO"]["cnpj_cpf"])
        dados.append(nota["DESTINATARIO"]["razao_social"])
        try:
            valor = nota["impostos"]["valor_inss"].replace('.', '').replace('R$ ', '')
        except:
            valor = '0,00'
        dados.append(valor)

        escreve_relatorio_csv(';'.join(dados))


def intervalo_comp(dti, dtf):
    dti = date(int(dti[1]), int(dti[0]), 1)
    dtf = date(int(dtf[1]), int(dtf[0]), 1)

    intervalo = []
    while True:
        tupla_comp = (str(dti.month).rjust(2, '0'), str(dti.year))
        intervalo.append(tupla_comp)
        
        if dti != dtf:
            dti = dti + relativedelta(months=+1)
        else:
            break

    return intervalo


def download_xml(nome, cnpj, senha, dtInicio, dtFinal):
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    url_pesquisa = "https://valinhos.sigissweb.com/nfecentral?oper=efetuapesquisasimples"
    url_xml = "https://valinhos.sigissweb.com/nfecentral?oper=geraxml&cnpj="

    # inicia a sessão no site, realiza o login e obtem os arquivos para o contribuinte
    with Session() as s:
        login_data = {"loginacesso":cnpj,"senha":senha}
        pagina = s.post(url_acesso, login_data)
        if pagina.status_code != 200:
            print('>>> Erro a acessar página.')
            return False

        for comp in intervalo_comp(dtInicio, dtFinal):
            info = {
                'cnpj_cpf_destinatario': '',
                'operCNPJCPFdest': 'EX',
                'RAZAO_SOCIAL_DESTINATARIO': '',
                'selnomedestoper': 'EX',
                'id_codigo_servico': '',
                'serie': '',
                'numero_nf': '',
                'operNFE': '=',
                'numero_nf2': '',
                'rps': '',
                'operRPS': '=',
                'rps2': '',
                'data_emissao':'', 
                'operData': '=',
                'data_emissao2': '',
                'mesi': comp[0],
                'anoi': comp[1],
                'mesf': comp[0],
                'anof': comp[1],
                'aliq_iss': '',
                'regime': '?',
                'iss_retido': '?',
                'cancelada': '?',
                'tipoPesq': 'normal'
            }

            pagina = s.post(url_pesquisa, info)
            if pagina.status_code != 200:
                print('>>> Erro a acessar página.')
                return False
            salvar = s.get(url_xml+cnpj)
            filename = salvar.headers.get('Content-Type', 'text')

            # rotina para salvar os arquivos pdf
            if 'text' not in filename:
                file = f'{nome.strip()} - {comp[0]}{comp[1]}.xml'
                arquivo = open(os.path.join('documentos', file), 'wb')
                for parte in salvar.iter_content(100000):
                    arquivo.write(parte)
                arquivo.close()
                print(f"Arquivo {file} salvo")

    return True


if __name__ == '__main__':
    comeco = datetime.now()
    os.makedirs('documentos', exist_ok=True)

    dtInicio = prompt("Data inicio no formato 00-0000:").split('-')
    dtFinal = prompt("Data inicio no formato 00-0000:").split('-')

    for empresa in empresas:
        print('>>> Buscando empresa', empresa[1])
        try:
            download_xml(*empresa, dtInicio, dtFinal)
        except Exception as e:
            print(e)

    for arq in os.listdir('documentos'):
        if arq[-4:] in ['.xml']:
            #print(arq)
            analisa_xml(arq)
    fim = datetime.now()
    print("\n>>> Tempo total:", (fim-comeco))
