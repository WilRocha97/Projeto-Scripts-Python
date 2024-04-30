# -*- coding: utf-8 -*-
from requests import Session
from bs4 import BeautifulSoup
import sys
import pyautogui as p
import re

sys.path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _time_execution, _open_lista_dados, _where_to_start, _indice


def lancamento(empresas, index):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)

        # cnpjform, cnpj, senha, nome = qual_empresa
        cnpjform, senha, nome, classif, numero, data, cnpjt, nomet, valor = empresa

        cnpjt = cnpjt.replace('.', '').replace('/', '').replace('-', '')
        cnpj = cnpjform.replace('.', '').replace('/', '').replace('-', '')

        periodo = data.split('/')

        with Session() as s:
            # entra no site
            s.get('https://valinhos.sigissweb.com/')

            # loga na empresa
            query = {'loginacesso': cnpj,
                     'senha': senha}
            s.post('https://valinhos.sigissweb.com/ControleDeAcesso', data=query)

            # entra na tela para lançar a nota
            s.get('https://valinhos.sigissweb.com/lancamentocentral?oper=novo')

            # lança a nota
            query = {'mes': periodo[1],
                     'ano': periodo[2],
                     'tomador_prestador': 'P',
                     'movimento': 'S',
                     'regime': 'S',
                     'cnpj_cpf_decl': cnpjform,
                     'regimeempresa': 'S',
                     'edomunicipio': 'true',
                     'cnpj_cpf_dest': cnpjt,
                     'classif': classif,
                     'documento_fiscal_canc': 'N',
                     'iss_retido_fonte': 'N',
                     'num_docu_fiscal': numero,
                     'serie_docu_fiscal': '',
                     'data': data,
                     'valor_docu_fiscal': valor,
                     'deducoes_legais': 'R$ 0,00',
                     'valor_servicos': valor,
                     'aliquota': '',
                     'valor_imposto': 'R$ 0,00',
                     'id_lancamento': 0,
                     'operacao': 'inssalvar',
                     'id_dest': 0,
                     'somenteLeitura': 'false',
                     'naopodealteraraliquota': 'false',
                     'tipoinclusao': 'M',
                     'btnsalvar': 'Salvar',
                     'btncancelar': 'Cancelar',
                     'estrangeiro': 0,
                     'oper': 'consistirEGravar',
                     'tipoInsert': 'M'}
            response = s.post('https://valinhos.sigissweb.com/lancamentocentral?', data=query)

            # analiza o código do site para pegar a resposta se a nota foi lançada ou não
            soup = BeautifulSoup(response.content, 'html.parser')
            soup = soup.prettify()
            padraozinho = re.compile(r'<titulo>\n. +(.+)\n.+\n.+\n.+<!\[CDATA\[(.+)]]>')
            matchzinho = padraozinho.search(soup)
            try:
                resposta = matchzinho.group(1) + ': ' + matchzinho.group(2)
                resposta = resposta.replace('<br>', ', ').replace('::', ':').replace('persisteLancamento - ', '')
            except:
                padraozinho = re.compile(r'<msgok>\n. +(.+)')
                matchzinho = padraozinho.search(soup)
                resposta = matchzinho.group(1)
                resposta = str(resposta)

            _escreve_relatorio_csv(';'.join([resposta, nome, data, numero, 'R$' + valor, cnpjt, nomet]), nome=f'Lançamento Tomador - Prestador {nome}')
            print('▶', resposta + '\n', nome, '/', numero)


def alterar(empresas, index):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)

        cnpjform, senha, nome, classif, periodo, num_nota, cnpj_cpf_dest, data_emissao, valor_nf, deducoes, valor_serv, aliquota, valor_imp = empresa

        cnpj = cnpjform.replace('.', '').replace('/', '').replace('-', '')

        with Session() as s:
            # entra no site
            s.get('https://valinhos.sigissweb.com/')

            # loga na empresa
            query = {'loginacesso': cnpj,
                     'senha': senha}
            s.post('https://valinhos.sigissweb.com/ControleDeAcesso', data=query)

            query = {'tomador_prestador': 'P',
                     'cnpj_cpf_decl': '',
                     'selcnpjcpfoperdecl': 'EX',
                     'cnpj_cpf_dest': '',
                     'selcnpjcpfoperdest': 'EX',
                     'cidade_dest': '',
                     'selcidadedestoperdest': 'EX',
                     'mesi': '01',
                     'anoi': '2018',
                     'mesf': '12',
                     'anof': '2021',
                     'mesicaixa': '',
                     'anoicaixa': '',
                     'mesfcaixa': '',
                     'anofcaixa': '',
                     'iss_retido_fonte': '?',
                     'regime': '?',
                     'movimento': '?',
                     'documento_fiscal_canc': '?',
                     'num_docu_fiscal_pesq': '',
                     'serie_docu_fiscal': '',
                     'data': '',
                     'aliquota': '?',
                     'caixaindicado': '?',
                     'classif': '?',
                     'lancamento_exclusos': '?'}

            response = s.post('https://valinhos.sigissweb.com/lancamentocentral?oper=efetuapesquisasimples', data=query)
            soup = BeautifulSoup(response.content, 'html.parser')
            soup = soup.prettify()
            padraozinho = re.compile(r'1 de (.+)\n.+</span>')
            match = padraozinho.search(soup)
            paginas = match.group(1)

            for i in range(int(paginas)):
                response = s.get(r'https://valinhos.sigissweb.com/lancamentocentral?oper=listarlancamentos&paginacao=3&numpag={}'.format(i))
                soup = BeautifulSoup(response.content, 'html.parser')
                soup = soup.prettify()
                padraozinho = re.compile(r"alteralancamento.+id_lancamento=(.+)&amp;resp=(.+)&amp;id_dest=(.+)' .+.\n.+ (" + num_nota + ")")
                match = padraozinho.search(soup)
                if match:
                    break
            try:
                id_lancamento = match.group(1)
            except:
                _escreve_relatorio_csv(';'.join(['Erro ao localizar a nota', num_nota, str(periodo)]), nome='Alteração da classificação de serviço')
                continue
            resp = match.group(2)
            id_dest = match.group(3)

            s.get(r'https://valinhos.sigissweb.com/lancamentocentral?oper=alterar&id_lancamento={}&resp={}&id_dest={}'.format(id_lancamento, resp, id_dest))

            periodo = periodo.split('/')
            query = {'mes': periodo[0],
                     'ano': periodo[1],
                     'tomador_prestador': 'P',
                     'movimento': 'S',
                     'regime': 'S',
                     'cnpj_cpf_decl': resp,
                     '51296390000120': cnpj,
                     'regimeempresa': 'S',
                     'edomunicipio': 'true',
                     'cnpj_cpf_dest': cnpj_cpf_dest,
                     'classif': classif,
                     'documento_fiscal_canc': 'N',
                     'iss_retido_fonte': 'N',
                     'num_docu_fiscal': num_nota,
                     'serie_docu_fiscal': '',
                     'data': data_emissao,
                     'valor_docu_fiscal': valor_nf,
                     'deducoes_legais': deducoes,
                     'valor_servicos': valor_serv,
                     'aliquota': aliquota,
                     'valor_imposto': valor_imp,
                     'id_lancamento': id_lancamento,
                     'operacao': 'altsalvar',
                     'id_dest': id_dest,
                     'somenteLeitura': 'false',
                     'naopodealteraraliquota': 'false',
                     'tipoinclusao': 'M',
                     'btnsalvar': 'Salvar',
                     'btncancelar': 'Cancelar',
                     'estrangeiro': '0',
                     'oper': 'consistirEGravar',
                     'tipoInsert': 'M'}

            response = s.post(r'https://valinhos.sigissweb.com/lancamentocentral?', data=query)
            soup = BeautifulSoup(response.content, 'html.parser')
            soup = soup.prettify()
            padrao = re.compile(r'<msgok>\n. +(.+)')
            match = padrao.search(soup)

            if not match:
                padrao = re.compile(r'<msg>\n. +(.+)')
                match = padrao.search(soup)

            resposta = match.group(1)
            resposta = resposta.replace('<![CDATA[', '').replace('.]]>', '')

            _escreve_relatorio_csv(';'.join([resposta, num_nota, str(periodo)]), nome='Alteração da classificação de serviço')
            print(resposta)


@_time_execution
def run():
    qual = p.confirm(title='Script incrível', text='Qual procedimento?', buttons=['Lançar nota', 'Alterar nota'])

    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    if qual == 'Lançar nota':
        lancamento(empresas, index)
    elif qual == 'Alterar nota':
        alterar(empresas, index)


if __name__ == '__main__':
    run()
