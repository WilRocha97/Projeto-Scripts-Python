# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
from dateutil.relativedelta import *
from datetime import datetime, date
from requests import Session
from pyautogui import prompt
import fitz, os, re, time

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice

'''
Download dos pdf no período dtInicio-dats_final para a empresa CNPJ (notas prestadas e tomadas)
Os arquivos são obtidos pelo site 'https://valinhos.sigissweb.com' através do usuario e senha 
do contribuinte.

url_acesso -> link para acessar o sistema
url_pesquisa -> link para filtrar o período necessário
url_sintético -> link que retorna o arquivo .pdf sintético do período solicitado
url_analitico -> link que retorna o arquivo .pdf analítico do período solicitado
'''


def consulta_xml(execucao, final, empresa, data_inicio, data_final):
    cnpj, senha, nome = empresa
    nome = nome.replace('/', ' ').replace('   ', ' ').replace('  ', ' ')
    
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    url_pesquisa = "https://valinhos.sigissweb.com/nfecentral?oper=efetuapesquisasimples"
    url_sintetico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfesintetico"
    url_analitico = "https://valinhos.sigissweb.com/nfecentral?oper=relnfeanalitico"
    # obter login e senha do banco de dados utilizando o cnpj

    # inicia a sessão no site, realiza o login e obtém os arquivos para o contribuinte
    with Session() as s:
        login_data = {"loginacesso": cnpj, "senha": senha}
        pagina = s.post(url_acesso, login_data)

        try:
            soup = BeautifulSoup(pagina.content, 'html.parser')
            soup = soup.prettify()
            # print(soup)
            regex = re.compile(r"'Aviso', '(.+)<br>")
            regex2 = re.compile(r"'Aviso', '(.+)\.\.\.', ")
            try:
                documento = regex.search(soup).group(1)
            except:
                documento = regex2.search(soup).group(1)
            _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}', local=final, nome=execucao)
            print(f"❌ {documento}")
            return False
        except:
            pass
    
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
            'data_emissao': '',
            'operData': '=',
            'data_emissao2': '',
            'mesi': data_inicio[0],
            'anoi': data_inicio[1],
            'mesf': data_final[0],
            'anof': data_final[1],
            'aliq_iss': '',
            'regime': '?',
            'iss_retido': '?',
            'cancelada': '?',
            'tipoPesq': 'normal'
        }

        pagina = s.post(url_pesquisa, info)
        if pagina.status_code != 200:
            _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Erro ao acessar a página', local=final, nome=execucao)
            print('>>> Erro ao acessar a página.')
            return False
        
        salvar = [(url_sintetico, os.path.join(final, 'Sintético'), 'Sintético'),
                  (url_analitico, os.path.join(final, 'Analítico'), 'Analítico')]
        
        # salvar relatório sintético depois relatório analítico
        for consulta in salvar:
            url = s.get(str(consulta[0]))
            pasta = consulta[1]
            if data_inicio != data_final:
                file = f'{nome} - {cnpj} - {consulta[2]} - {str(data_inicio[0])}-{str(data_inicio[1])} até {str(data_final[0])}-{str(data_final[1])}'
            else:
                file = f'{nome} - {cnpj} - {consulta[2]} - {str(data_inicio[0])}-{str(data_inicio[1])}'
                
            # rotina para salvar os arquivos pdf
            if 'text' not in url.headers.get('Content-Type', 'text'):
                os.makedirs(pasta, exist_ok=True)
                os.makedirs(pasta, exist_ok=True)
                
                arquivo = open(os.path.join(pasta, file + '.pdf'), 'wb')
                for parte in url.iter_content(100000):
                    arquivo.write(parte)
                arquivo.close()
                
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Relatório {consulta[2]} gerado', local=final, nome=execucao)
                print(f"✔ Relatório {consulta[2]} salvo")
            else:
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};Erro ao acessar a página {consulta[2]}', local=final, nome=execucao)
                print(f'❌ Não gerou relatório {consulta[2]}.')

            # necessário para não sobrepor o cachê da pesquisa
            time.sleep(1)
                
    return True


def analiza(final):
    print('\n>>> Analisando relatórios Analíticos')
    documentos = os.path.join(final, 'Analítico')
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(documentos):
        # Abrir o pdf
        arquivo = os.path.join(documentos, arq)
        with fitz.open(arquivo) as pdf:
            
            pattern = re.compile(r' - \d+ - Analítico.+.pdf')
            empresa = re.sub(pattern, '', arq)
            
            cnpj = re.compile(r' - (\d+) - Analítico.+.pdf').search(arq).group(1)
            
            # Para cada página do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    # Procura o valor a recolher da empresa
                    valores = None
                    indice = 20
                    while not valores:
                        indice = str(indice)
                        valores = re.compile(r'\n(.+)\n.+\n.0+(.+)\n(.+)\n(.+)\nSérie:\n(.+\n){' + indice + '}ISS\n.+\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)\nR\$ (.+)(\n.+){7}\nR\$ (.+)').findall(textinho)
                        indice = int(indice)
                        indice += 1
                        
                        if indice >= 30:
                            break
                    
                    if not valores:
                        continue
                    
                    for valor in valores:
                        # Guarda as infos da empresa
                        emissao = valor[0]
                        numero = valor[1]
                        cnpj_destinatario = valor[2]
                        nome_destinatario = valor[3]
                        iss = valor[11]
                        pis = valor[5]
                        cofins = valor[8]
                        csll = valor[9]
                        inss = valor[6]
                        irrf = valor[7]
                        nome_destinatario = nome_destinatario.replace('–', '-')
                        _escreve_relatorio_csv(f"{cnpj};{empresa};{emissao};{numero};{nome_destinatario};{cnpj_destinatario};{iss};{pis};{cofins};{csll};{inss};{irrf}",
                                               nome='Resumo Relatórios Analíticos', local=final)
                except():
                    print(f'\nArquivo: {arq} - ERRO')


@_time_execution
def run():
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    execucao = 'Resumo Download Relatórios'
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    data_inicio = prompt("Data inicio no formato 00/0000:").split('/')
    data_final = prompt("Data final no formato 00/0000:").split('/')
    if data_inicio != data_final:
        final = os.path.join('Execução', 'Ref. ' + data_inicio[0] + '-' + data_inicio[1] + ' até ' + data_final[0] + '-' + data_final[1])
    else:
        final = os.path.join('Execução', 'Ref. ' + data_inicio[0] + '-' + data_inicio[1])
        
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        try:
            consulta_xml(execucao, final, empresa, data_inicio, data_final)
        except Exception as e:
            print(e)
    
    time.sleep(1)
    analiza(final)
    _escreve_header_csv(';'.join(['CNPJ', 'NOME', 'EMISSÃO DA NOTA', 'NÚMERO DA NOTA', 'DESTINATÁRIO', 'CNPJ DESTINATÁRIO', 'ISS', 'PIS', 'COFINS', 'CSLL', 'INSS', 'IRRF']),
                        nome='Resumo Relatórios Analíticos', local=final)

if __name__ == '__main__':
    run()
