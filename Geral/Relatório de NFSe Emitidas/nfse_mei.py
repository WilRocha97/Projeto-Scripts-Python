# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, re, fitz, shutil

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, usuario, senha):
    print('>>> Logando no site')
    timer = 0
    while not _find_by_id('Inscricao', driver):
        try:
            driver.get('https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional')
        except:
            print('>>> Site demorou pra responder, tentando novamente')
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'
    
    # insere o usuário e a senha no site
    _send_input('Inscricao', usuario, driver)
    _send_input('Senha', senha, driver)
    time.sleep(1)
    
    # clica no botão para logar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button').click()
    time.sleep(1)
    return driver, 'ok'

    
def consulta_notas(download_folder, driver, cnpj, nome):
    alerta = re.compile(r'<div class=\"alert-warning alert\"><span class=\"icone\"></span>(.+)<a class=\"close\"').search(driver.page_source)
    if alerta:
        return driver, f'❗ {alerta.group(1)}'

    if re.compile(r'Até o momento nenhuma NFS-e foi emitida').search(driver.page_source):
        return driver, '❗ Até o momento nenhuma NFS-e foi emitida'
    
    print('>>> Consultando notas emitidas')
    driver.get('https://www.nfse.gov.br/EmissorNacional/Notas/Emitidas')
    time.sleep(1)
    # captura os links para entrar em cada nota da lista
    link_notas = re.compile(r'href=\"(/EmissorNacional/Notas/Visualizar/Index/.+)\" class=\"list-group-item\">').findall(driver.page_source)
    
    if not link_notas:
        print(driver.page_source)
    
    # para cada nota da lista
    dados = []
    for count, link_nota in enumerate(link_notas, start=1):
        print(f'>>> Abrindo {count}° nota')
        # abre a nota
        driver.get('https://www.nfse.gov.br' + link_nota)
        time.sleep(1)
        
        # coleta os dados da nota no site
        dados_do_site = coleta_dados_da_nota_no_site(driver)
        
        # verifica se a nota foi cancelada
        nf_cancelada = ' - '
        situacao = '0'
        if re.compile(r'(Evento de Cancelamento de NFS-e)').search(driver.page_source):
            nf_cancelada = ' - CANCELADA - '
            situacao = '2'
            
        # pega o link para baixar o PDF da nota
        link_pdf = re.compile(r'href=\"(/EmissorNacional/Notas/Download/DANFSe/.+)\" class=\"btn btn-lg btn-info\"').search(driver.page_source).group(1)
        
        # clica no botão para download da NFSE
        driver.get('https://www.nfse.gov.br' + link_pdf)
        time.sleep(1)

        dados_pdf = mover_arquivo(download_folder, link_pdf, cnpj, nome, nf_cancelada)
        
        dados.append(f'{dados_do_site};{situacao};{dados_pdf}')
        
    for nota in dados:
        _escreve_relatorio_csv(nota, nome='Notas')
        
    return driver, 'ok'


def coleta_dados_da_nota_no_site(driver):
    print('>>> Analisando dados no site')
    # lista com os regex para pegar cada item da nota
    print(driver.page_source)
    time.sleep(33)
    
    cnpj_tomador = re.compile(r'Tomador.+(\n.+){4}CNPJ</span></label><span class=\"form-control-static .+\"(.+)</span></div>').search(driver.page_source).group(2)
    nome_tomador = re.compile(r'Tomador.+(\n.+){14}Razão Social</span></label><span class=\"form-control-static .+\"(.+)</span></div>').search(driver.page_source).group(2)
    endereco_tomador = re.compile(r'texto\">(.+\n.+\n.+)(\n.+){3},\s+(.+)</span></div>(\s+</div>\n){5}\s+<div id=\"servicos\" class=\"tab-pane fade\">').search(driver.page_source)
    
    uf = endereco_tomador.group(3).split('/')[1]
    cidade = endereco_tomador.group(3).split('/')[0]
    endereco = endereco_tomador.group(1).re.sub(r'\s+', '').replace('\n', ' ')
    
    numero_nota = re.compile(r'Número</span></label><span class=\"form-control-static texto\">(.+)</span></div>').search(driver.page_source).group(1)
    serie_nota = re.compile(r'Série</span></label><span class=\"form-control-static texto\">(.+)</span></div>').search(driver.page_source).group(1)
    data_emissao = re.compile(r'Data de emissão</span></label><span class=\"form-control-static texto\">(.+) \n\s+ .+\n\s+ .+</span></div>').search(driver.page_source).group(1)
    
    return f'{cnpj_tomador};{nome_tomador};{uf};{cidade};{endereco};{numero_nota};{serie_nota};{data_emissao}'


def mover_arquivo(download_folder, link_pdf, cnpj, nome, nf_cancelada):
    nome_arquivo = link_pdf.split('/')[-1] + '.pdf'
    
    for arquivo in os.listdir(download_folder):
        while re.compile(r'crdownload').search(arquivo):
            print('>>> Aguardando download...')
            time.sleep(3)
    
    print('>>> Analisando PDF')
    # abre o PDF da nota para capturar algumas infos para adicionar na planilha de andamentos e para renomear o arquivo
    with fitz.open(os.path.join(download_folder, nome_arquivo)) as pdf:
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            
            # captura dados para o nome do arquivo
            numero_nf = re.compile(r'NúmerodaNFS-e\n(.+)').search(textinho).group(1)
            competencia = re.compile(r'CompetênciadaNFS-e\n(.+)').search(textinho).group(1)
            tomador = re.compile('TOMADORDOSERVIÇO\nCNPJ/CPF/NIF\n(.+)').search(textinho)
            if not tomador:
                tomador = re.compile('TOMADORDOSERVIÇONÃOIDENTIFICADONANFS-e').search(textinho)
                if tomador:
                    tomador = 'Tomador não identificado na NFS-e'
            else:
                tomador = tomador.group(1)
            tomador = tomador.replace('.', '').replace('/', '').replace('-', '')
            novo_arquivo = f'{nome} - NFSE-{numero_nf} - {competencia.replace("/", "-")}{nf_cancelada}Tomador_{tomador}.pdf'
            
            # captura dados para a planilha
            acumulador = '6102'
            local_do_servico = re.compile(r'LocaldaPrestação\n(.+)').search(textinho).group(1)
            if local_do_servico == 'Valinhos-SP':
                cfps = '9101'
            else:
                cfps = '9102'
            valor_servico = re.compile(r'ValordoServiço\nR\$(.+)').search(textinho).group(1)
            valor_descontos = '0,00'
            valor_deducoes = re.compile(r'TotalDeduções/Reduções\n(.+)').search(textinho).group(1)
            if valor_deducoes == '-':
                valor_deducoes = '0,00'
            valor_contabil = valor_servico
            base_calculo = valor_servico
            aliquota_iss = '0,00'
            valor_iss = '0,00'
            valor_iss_retido = '0,00'
            valor_irrf = '0,00'
            valor_pis = '0,00'
            valor_cofins = '0,00'
            valor_csll = '0,00'
            
            dados_pdf = (f'{acumulador};'
                         f'{cfps};'
                         f'{valor_servico};'
                         f'{valor_descontos};'
                         f'{valor_deducoes};'
                         f'{valor_contabil};'
                         f'{base_calculo};'
                         f'{aliquota_iss};'
                         f'{valor_iss};'
                         f'{valor_iss_retido};'
                         f'{valor_irrf};'
                         f'{valor_pis};'
                         f'{valor_cofins};'
                         f'{valor_csll}')
            
    # move e renomeio o arquivo
    shutil.move(os.path.join(download_folder, nome_arquivo), os.path.join(download_folder, novo_arquivo))
    print(novo_arquivo)
    
    return dados_pdf


@_time_execution
def run():
    download_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\Relatório de NFSe Emitidas\\Execução\\NFSE"
    os.makedirs(download_folder, exist_ok=True)
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,  # muda o diretório padrão de download do navegador
        "download.prompt_for_download": False,  # faz o download automatico sem perguntar onde salvar
        "download.directory_upgrade": True,  # atualiza o diretório de download padrão do navegador
        "plugins.always_open_pdf_externally": True  # não irá abrir o PDF no navegador
    })
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, usuario, senha = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            driver, situacao = login(driver, usuario, senha)
            if situacao == 'ok':
                driver, resultado = consulta_notas(download_folder, driver, cnpj, nome)
                if resultado != 'ok':
                    print(resultado)
                    _escreve_relatorio_csv(f'{cnpj};{nome};{resultado[2:]}', nome='Ocorrências')
                break
                
            driver.close()
        driver.close()
    
    # escreve o cabeçalho da planilha
    _escreve_header_csv('CNPJ;'
                        'NOME;'
                        'NÚMERO DA NF;'
                        'NÚMERO DA DPS;'
                        'CHAVE DE ACESSO;'
                        'COMPETÊNCIA;'
                        'DATA DE EMISSÃO;'
                        'AUTOR DO EVENTO DE CANCELAMENTO;'
                        'DATA DO EVENTO;'
                        'DATA DE REGISTRO DO EVENTO;'
                        'VERSÃO;'
                        'SÉRIE;'
                        'CNPJ EMITENTE;'
                        'RAZÃO SOCIAL EMITENTE;'
                        'INSCRIÇÃO MUNICIPAL;'
                        'SITUAÇÃO PERANTE O SIMPLES NACIONAL;'
                        'REGIME ESPECIAL DE TRIBUTAÇÃO;'
                        'EMITENTE EMAIL;'
                        'CNPJ TOMADOR;'
                        'RAZÃO SOCIAL TOMADOR;'
                        'INSCRIÇÃO MUNICIPAL TOMADOR;'
                        'TRIBUTAÇÃO ISSQN;'
                        'PAÍS DA PRESTAÇÃO DO SERVIÇO;'
                        'MUNICÍPIO DE INCIDÊNCIA;'
                        'TIPO DE IMUNIDADE;'
                        'SUSPENSÃO DO ISSQN;'
                        'NÚMERO DO PROCESSO DE SUSPENSÃO;'
                        'BENEFÍCIO MUNICIPAL;'
                        'VALOR DO SERVIÇO;'
                        'DESCONTO INCONDICIONADO;'
                        'TOTAL DE DEDUÇÕES/REDUÇÕES;'
                        'TOTAL DO BENEFÍCIO MUNICIPAL;'
                        'BASE DE CÁLCULO;'
                        'ALÍQUOTA;'
                        'VALOR DO ISSQN;'
                        'RETENÇÃO;'
                        'VERSÃO DA APLICAÇÃO;'
                        'AMBIENTE GERADOR;'
                        'SITUAÇÃO DA NFS-E;'
                        'PAÍS;'
                        'MUNICÍPIO;'
                        'CÓD. DE TRIBUTAÇÃO NACIONAL;'
                        'ITEM DA NBS CORRESP. AO SERVIÇO PRESTADO;'
                        'DESCRIÇÃO DO SERVIÇO;'
                        'NOME DO EVENTO;'
                        'DATA DE INÍCIO;'
                        'DATA DO FIM;'
                        'CÓDIGO DE IDENTIFICAÇÃO;'
                        'Nº DOC. RESPONSABILIDADE TÉCNICA;'
                        'DOC. REFERÊNCIA;'
                        'INFOS COMPLEMENTARES;'
                        'SITUAÇÃO TRIB. PIS/COFINS;'
                        'BASE DE CÁLCULO PIS/COFINS;'
                        'PIS-ALÍQUOTA;'
                        'PIS-VALOR IMPOSTO;'
                        'COFINS-ALÍQUOTA;'
                        'COFINS-VALOR IMPOSTO;'
                        'TIPO RETENÇÃO PIS/COFINS;'
                        'VALOR RETIDO IRRF;'
                        'VALOR RETIDO CSLL;V'
                        'ALOR RETIDO CP;'
                        'OPÇÃO')
    
    
if __name__ == '__main__':
    run()
