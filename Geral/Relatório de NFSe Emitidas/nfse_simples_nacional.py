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
    print('>>> Consultando notas emitidas')
    driver.get('https://www.nfse.gov.br/EmissorNacional/Notas/Emitidas')
    time.sleep(1)
    # captura os links para entrar em cada nota da lista
    link_notas = re.compile(r'href=\"(/EmissorNacional/Notas/Visualizar/Index/.+)\" class=\"list-group-item\">').findall(driver.page_source)
    
    if not link_notas:
        print(driver.page_source)
    
    # para cada nota da lista
    for count, link_nota in enumerate(link_notas, start=1):
        print(f'>>> Abrindo {count}° nota')
        # abre a nota
        driver.get('https://www.nfse.gov.br' + link_nota)
        time.sleep(1)
        # coleta os dados da nota no site
        dado_do_site = coleta_dados_da_nota(driver)
        
        # verifica se a nota foi cancelada
        nf_cancelada = ''
        if re.compile(r'(Evento de Cancelamento de NFS-e)').search(driver.page_source):
            nf_cancelada = ' - CANCELADA'
            
        # pega o link para baixar o PDF da nota
        link_pdf = re.compile(r'href=\"(/EmissorNacional/Notas/Download/DANFSe/.+)\" class=\"btn btn-lg btn-info\"').search(driver.page_source).group(1)
        # clica no botão para download da NFSE
        driver.get('https://www.nfse.gov.br' + link_pdf)
        time.sleep(1)
        # captura dados do PDF da nota
        dados_pdf = mover_arquivo(download_folder, link_pdf, cnpj, nome, nf_cancelada)
        
        _escreve_relatorio_csv(f'{dados_pdf};{dado_do_site}')
        
    return driver


def coleta_dados_da_nota(driver):
    print('>>> Analisando dados no site')
    # lista com os regex para pegar cada item da nota
    dados = [r'(Data de emissão)</span></label><span class=\"form-control-static texto\">(.+) \n\s+( .+)\n\s+( .+)</span></div>',
             r'(Autor do Evento)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Data do Evento)</span></label><span class=\"form-control-static .+\">(.+) \n\s+( .+)\n\s+( .+)</span></div>',
             r'(Data do registro do Evento)</span></label><span class=\"form-control-static .+\">(.+) \n\s+( .+)\n\s+( .+)</span></div>',
             r'(Versão)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Série)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Emitente.+(\n.+){9}CNPJ</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Emitente.+(\n.+){4}Razão Social</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Emitente.+(\n.+){12}Inscrição Municipal</span></label><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Situação Perante o Simples Nacional)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Regime Especial de Tributação)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Emitente.+(\n.+){41}Email</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Tomador.+(\n.+){4}CNPJ</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Tomador.+(\n.+){14}Razão Social</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'Tomador.+(\n.+){7}Inscrição Municipal</span></label><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Tributação do ISSQN)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(País Resultado da Prestação de Serviço)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Município de Incidência)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Tipo de Imunidade)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Suspensão do ISSQN)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Número processo suspensão)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Benefício Municipal - BM)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Valor do Serviço.+R\$)</span><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Desconto incondicionado.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Total Deduções/Reduções.+R\$)</span><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Total Benefício Municipal.+R\$)</span><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Base de Cálculo.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Alíquota)</span></label><div class=\"input-group\"><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Valor do ISSQN.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Retenção)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Versão da Aplicação)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Ambiente Gerador)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Situação da NFS-e)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(País)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Município)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Código de Tributação Nacional)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Item da NBS correspondente ao serviço prestado)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Descrição do serviço)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Nome do Evento)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Data de início)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Data de fim)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Código de identificação)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Número de documento de responsabilidade técnica)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Documento de referência)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Informações complementares)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Situação tributária do PIS/COFINS)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Base de cálculo PIS/COFINS.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(PIS - Alíquota)</span></label><div class=\"input-group\"><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(PIS - Valor do imposto.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(COFINS - Alíquota)</span></label><div class=\"input-group\"><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(COFINS - Valor do imposto.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Tipo de Retenção do PIS/Cofins)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',
             r'(Valor Retido IRRF.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Valor Retido CSLL.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Valor Retido CP.+R\$)</span><span class=\"form-control-static .+\">\n\s+(.+)',
             r'(Opção)</span></label><span class=\"form-control-static .+\"(.+)</span></div>',]
    
    dado_do_site = ''
    dado_composto = ''
    # para cada regex, procura a info na nota
    for count, dado in enumerate(dados):
        # procura o dado na nota
        iten = re.compile(dado).search(driver.page_source)
        # se não encontrar o dado específico, coloca um "-" na variável
        if not iten:
            dado_do_site += '-;'
        else:
            for i in range(8):
                if i < 2:
                    continue
                try:
                    dado_composto += iten.group(i).replace('>', '').replace(' - ', '-')
                except:
                    pass
                    
            dado_do_site += dado_composto + ';'
            dado_composto = ''
        # adiciona todos os dados formatados com ";" para escrever na planilha
    
    return dado_do_site


def mover_arquivo(download_folder, link_pdf, cnpj, nome, nf_cancelada):
    nome_arquivo = link_pdf.split('/')[-1] + '.pdf'
    print('>>> Analisando PDF')
    # abre o PDF da nota para capturar algumas infos para adicionar na planilha de andamentos e para renomear o arquivo
    with fitz.open(os.path.join(download_folder, nome_arquivo)) as pdf:
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            
            numero_nf = re.compile(r'NúmerodaNFS-e\n(.+)').search(textinho).group(1)
            numero_dps = re.compile(r'NúmerodaDPS\n(.+)').search(textinho).group(1)
            competencia = re.compile(r'CompetênciadaNFS-e\n(.+)').search(textinho).group(1)
            chave_acesso = "'" + re.compile(r'ChavedeAcessodaNFS-e\n(.+)').search(textinho).group(1)
            tomador = re.compile('TOMADORDOSERVIÇO\nCNPJ/CPF/NIF\n(.+)').search(textinho).group(1)
            tomador = tomador.replace('.', '').replace('/', '').replace('-', '')
            
            dados_pdf = f'{cnpj};{nome};{numero_nf};{numero_dps};{chave_acesso};{competencia}'
            novo_arquivo = f'NFSE_{numero_nf} - {competencia.replace("/", "-")} - Prestador_{cnpj} - Tomador_{tomador}{nf_cancelada}.pdf'
    
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
                driver = consulta_notas(download_folder, driver, cnpj, nome)
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
