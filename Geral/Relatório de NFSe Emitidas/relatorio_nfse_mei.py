# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from decimal import Decimal
from pyautogui import confirm, prompt
import os, time, re, fitz, shutil
import datetime

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, usuario, senha):
    print('>>> Logando no site')
    # espera até 1 minuto o site carregar, se não carregar retorna e tenta de novo
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
    try:
        driver.find_element(by=By.XPATH, value='/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button').click()
    except:
        return driver, 'erro'
    
    time.sleep(1)
    return driver, 'ok'

    
def consulta_notas(andamentos, periodo, download_folder, driver, cod_dominio, cnpj, nome):
    # verifica se aparece alguma mensagem de alerta no site
    alerta = re.compile(r'<div class=\"alert-warning alert\"><span class=\"icone\"></span>(.+)<a class=\"close\"').search(driver.page_source)
    if alerta:
        return driver, f'❗ {alerta.group(1)}'
    
    # verifica se a empresa tem notas no site
    if re.compile(r'Até o momento nenhuma NFS-e foi emitida').search(driver.page_source):
        return driver, '❗ Até o momento nenhuma NFS-e foi emitida'
    
    print('>>> Consultando notas emitidas')
    # entra na página com a lista de notas emitidas
    paginas = 1
    quantidade_de_notas = 0
    dados = []
    dados_para_somar_valores = []
    while True:
        # percorre infinitas páginas
        driver.get('https://www.nfse.gov.br/EmissorNacional/Notas/Emitidas?pg=' + str(paginas) + '&=2')
        time.sleep(1)
        # se chegar em uma página onde não existir registros, sai do loop
        if re.compile(r'Nenhum registro encontrado').search(driver.page_source):
            break
        # captura os links para entrar em cada nota da lista
        link_notas = re.compile(r'href=\"(/EmissorNacional/Notas/Visualizar/Index/.+)\" class=\"list-group-item\">').findall(driver.page_source)
        
        # se não encontrar a lista de notas printa o código do site para análise
        if not link_notas:
            print(driver.page_source)
        
        # para cada nota da lista
        notas_no_periodo = 0
        for link_nota in link_notas:
            # abre a nota
            driver.get('https://www.nfse.gov.br' + link_nota)
            time.sleep(1)
            
            # coleta os dados da nota no site
            dados_do_site, data_emissao, nome_tomador = coleta_dados_da_nota_no_site(driver)
            periodo_da_nota = re.search(r'\d{2}/(\d{2}/\d{4})', data_emissao).group(1)
            
            if re.compile(r'\d\d/\d\d\d\d').search(periodo):
                if periodo != '' and periodo != periodo_da_nota:
                    print('>>> Essa nota não corresponde ao período pesquisado\n')
                    continue
            
            else:
                if periodo != '' and periodo != periodo_da_nota.split('/')[1]:
                    print('>>> Essa nota não corresponde ao período pesquisado\n')
                    continue
            
            notas_no_periodo += 1
            # cria a pasta para salvar as notas
            os.makedirs(download_folder, exist_ok=True)
            # guarda a quantidade de notas abertas
            quantidade_de_notas += 1
            print(f'\n>>> Abrindo {quantidade_de_notas}° nota')
            
            # verifica se a nota foi cancelada
            nf_cancelada = ' - '
            situacao = '0'
            if re.compile(r'(Evento de Cancelamento de NFS-e)').search(driver.page_source):
                nf_cancelada = ' - CANCELADA - '
                situacao = '2'
                
            # pega o link para baixar o PDF da nota
            link_pdf = re.compile(r'href=\"(/EmissorNacional/Notas/Download/DANFSe/.+)\" class=\"btn btn-lg btn-info\"').search(driver.page_source).group(1)
            
            # clica no botão para download da NFSE
            print('>>> Fazendo download da nota...')
            driver.get('https://www.nfse.gov.br' + link_pdf)
            time.sleep(1)
            
            # move para pasta final e captura informações para anotar na planilha e renomear o arquivo
            dados_pdf, valor_total = coleta_dados_e_renomeia_arquivo(download_folder, link_pdf, nome, nf_cancelada)
            
            # se ao coletar os dados no site for encontrado que o tomador não foi informado na nota, não irá anotar os dados na planilha referente a essa nota,
            # só cria txt das notas com tomador informado
            if nome_tomador != 'O tomador e o intermediário não foram identificados pelo emitente' and nome_tomador != 'Nota fiscal sem tomador, apenas intermediário':
                dados.append(f'{dados_do_site};{situacao};{dados_pdf}')
            # adiciona na planilha todas as notas
            _escreve_relatorio_csv(f'{cod_dominio};{cnpj};{nome};{dados_do_site};{situacao};{dados_pdf}', nome=andamentos)
            
            # cria uma lista só com as datas
            if situacao != '2':
                data_emissao = data_emissao.split('/')
                dados_para_somar_valores.append((f'{data_emissao[1]}/{data_emissao[2]}', valor_total))
        
        if notas_no_periodo < 0:
            return driver, 'ok'
        
        paginas += 1
    
    comp_nota_anterior = ''
    valor_total_periodo = 0
    contador = 0
    dados_para_somar_valores = sorted(dados_para_somar_valores)
    for count, dado in enumerate(dados_para_somar_valores):
        #if count == 0:
        valor_total_periodo = valor_total_periodo + Decimal(float(str(dado[1]).replace('.', '').replace(',', '.')))
        comp_nota_anterior = dado[0]
        contador += 1
        #continue
            
        """elif dado[0] == comp_nota_anterior:
            valor_total_periodo = valor_total_periodo + Decimal(float(str(dado[1]).replace('.', '').replace(',', '.')))
            comp_nota_anterior = dado[0]
            contador += 1
            continue
    
        valor_total_periodo = round(valor_total_periodo, 2)
        valor_total_periodo = str(valor_total_periodo).replace('.', ',')
        # anota na planilha de resumo o andamento da consulta e quantas notas foram baixadas
        _escreve_relatorio_csv(f'{cod_dominio};{cnpj};{nome};{contador} notas encontradas;{periodo};{valor_total_periodo}')
        
        valor_total_periodo = Decimal(float(str(dado[1]).replace('.', '').replace(',', '.')))
        comp_nota_anterior = dado[0]
        contador = 1"""
    
    valor_total_periodo = round(valor_total_periodo, 2)
    valor_total_periodo = str(valor_total_periodo).replace('.', ',')
    # anota na planilha de resumo o andamento da consulta e quantas notas foram baixadas
    _escreve_relatorio_csv(f'{cod_dominio};{cnpj};{nome};{contador} notas encontradas;{periodo};{valor_total_periodo}')
    
    # para cada nota armazenada na variável, insere na planilha de notas
    pula_linha_txt = ''
    for count, nota in enumerate(dados):
        if count > 0:
            pula_linha_txt = '\n'
        # cria um txt com os dados das notas em uma pasta nomeada com o código no domínio e o cnpj do prestador
        cria_txt(cod_dominio, cnpj, nota, pula_linha_txt)
    
    return driver, 'ok'


def coleta_dados_da_nota_no_site(driver):
    print('>>> Analisando dados no site')
    # print(driver.page_source)
    # time.sleep(3)
    
    # verifica se o tomador foi informado na nota
    # se não foi informado retorna e não coleta mais nada, pois só irá anotar na planilha se o tomador for informado
    nao_tem_tomador = re.compile(r'O tomador e o indermediário não foram identificados pelo emitente').search(driver.page_source)
    tem_intermediario = re.compile(r'Intermediário').search(driver.page_source)
    
    if nao_tem_tomador:
        nome_tomador = 'O tomador e o intermediário não foram identificados pelo emitente'
    
    elif tem_intermediario:
        nome_tomador = 'Nota fiscal sem tomador, apenas intermediário'
    
    else:
        # captura o nome do tomador, o campo pode variar de posição no código do site, por isso o 'for' para editar o regex até encontrar
        for i in range(20):
            try:
                nome_tomador = re.compile(r'Tomador.+(\n.+){' + str(i) + '}Razão Social</span></label><span class=\"form-control-static .+\">(.+)</span></div>').search(driver.page_source).group(2)
                nome_tomador = nome_tomador.replace('&amp;', '&')
                break
            except:
                pass
    
    data_emissao = re.compile(r'Data de emissão</span></label><span class=\"form-control-static texto\">(.+) \n\s+ .+\n\s+ .+</span></div>').search(driver.page_source).group(1)
    
    # captura o CNPJ ou CPF do tomador
    # a variável é iniciada fora do 'for', pois caso não encontre pode ser se o tomador seja fora do Brasil
    cpf_cnpj_tomador = ''
    for id_tomador in ['CNPJ', 'CPF']:
        try:
            cpf_cnpj_tomador = re.compile(r'Tomador.+(\n.+){4}' + id_tomador + '</span></label><span class=\"form-control-static .+\">(.+)</span></div>').search(driver.page_source).group(2)
            break
        except:
            pass
    
    # captura o endereço do tomador, duas variações
    enderecos_tomador = re.compile(r'Endereço do Estabelecimento/Domicílio.+texto\">(.+\n\s+.+\n\s+.+)(\n.+){3}\n\s+(.+)</span></div>').findall(driver.page_source)
    if not enderecos_tomador:
        enderecos_tomador = re.compile(r'Endereço do Estabelecimento/Domicílio.+texto\">(.+\n\s+.+)(\n.+){3}\n\s+(.+)</span></div>').findall(driver.page_source)
    
    # inicia as variáveis de 'uf', 'cidade' e 'endereco' fora do 'for', pois de for fora do Brasil, não precisa colocar na planilha de notas
    uf = ''
    cidade = ''
    endereco = ''
    for count, endereco_tomador in enumerate(enderecos_tomador):
        # pega apenas a última correspondência do regex, pois o endereço do tomador é o último que aparece na nota
        if count == len(enderecos_tomador) - 1:
            # tenta pegar as infos de endereço, se não conseguir provavelmente é endereço fora do Brasil, nesse caso pode deixar em branco
            try:
                uf = endereco_tomador[2].split('/')[1]
                cidade = endereco_tomador[2].split('/')[0]
                endereco = endereco_tomador[0]
                endereco = re.sub(r'  +', '', endereco).replace('\n', ' ').replace('\r', ' ').replace(' , ', ', ')
            except:
                pass
    
    # pega mais informações da nota
    numero_nota = re.compile(r'Número</span></label><span class=\"form-control-static texto\">(.+)</span></div>').search(driver.page_source).group(1)
    serie_nota = re.compile(r'Série</span></label><span class=\"form-control-static texto\">(.+)</span></div>').search(driver.page_source).group(1)
    
    return f'{cpf_cnpj_tomador};{nome_tomador};{uf};{cidade};{endereco};{numero_nota};{serie_nota};{data_emissao}', data_emissao, nome_tomador


def coleta_dados_e_renomeia_arquivo(download_folder, link_pdf, nome, nf_cancelada):
    nome_arquivo = link_pdf.split('/')[-1] + '.pdf'
    
    # verifica se terminou o download do arquivo
    for arquivo in os.listdir(download_folder):
        while re.compile(r'crdownload').search(arquivo):
            print('>>> Aguardando download...')
            time.sleep(3)
            for arq in os.listdir(download_folder):
                arquivo = arq
    
    print('>>> Analisando PDF')
    # abre o PDF da nota para capturar algumas infos para adicionar na planilha de andamentos e para renomear o arquivo
    while True:
        try:
            with fitz.open(os.path.join(download_folder, nome_arquivo)) as pdf:
                break
        except:
            pass
        
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
                    print(textinho)
                    time.sleep(33)
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
            
            irrf_cp_csll_retido = re.compile(r'IRRF,CP,CSLL-Retidos\nR\$(.+)').search(textinho).group(1)
            pis_cofins_retido = re.compile(r'PIS/COFINSRetidos\n(.+)').search(textinho).group(1)
            if pis_cofins_retido == '-':
                pis_cofins_retido = '0,00'
            valor_crf = float(irrf_cp_csll_retido.replace(',', '.')) + float(pis_cofins_retido.replace(',', '.'))
            
            valor_inss = '0,00'
            codigo_item = re.compile(r'CódigodeTributaçãoNacional\n(.+)-').search(textinho).group(1)
            quantidade = '1,00'
            valor_unitario = valor_servico
            
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
                         f'{valor_csll};'
                         f'{valor_crf};'
                         f'{valor_inss};'
                         f'{codigo_item};'
                         f'{quantidade};'
                         f'{valor_unitario}')
            
    # renomeio o arquivo
    shutil.move(os.path.join(download_folder, nome_arquivo), os.path.join(download_folder, novo_arquivo))
    print(novo_arquivo)
    
    return dados_pdf.replace('.', ''), valor_servico


def cria_txt(codigo, cnpj, dados_nota, pula_linha_txt, tipo='NFSe-MEI', encode='latin-1'):
    # cria um txt com os dados das notas em uma pasta nomeada com o código no domínio e o cnpj do prestador
    local = os.path.join('Execução', 'Arquivos para Importação', str(codigo) + '-' + str(cnpj))
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{tipo} - {cnpj}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{tipo}  - {cnpj} - auxiliar.txt"), 'a', encoding=encode)
    """ ESTRUTURA DOS DADOS DA NOTA """
    """dados_site = (f'{cpf_cnpj_tomador};' # 0
                    f'{nome_tomador};'      # 1
                    f'{uf};'                # 2
                    f'{cidade};'            # 3
                    f'{endereco};'          # 4
                    f'{numero_nota};'       # 5
                    f'{serie_nota};'        # 6
                    f'{data_emissao}')"""   # 7
    """ situacao """                        # 8
    """dados_pdf = (f'{acumulador};'        # 9
                 f'{cfps};'                 # 10
                 f'{valor_servico};'        # 11
                 f'{valor_descontos};'      # 12
                 f'{valor_deducoes};'       # 13
                 f'{valor_contabil};'       # 14
                 f'{base_calculo};'         # 15
                 f'{aliquota_iss};'         # 16
                 f'{valor_iss};'            # 17
                 f'{valor_iss_retido};'     # 18
                 f'{valor_irrf};'           # 19
                 f'{valor_pis};'            # 20
                 f'{valor_cofins};'         # 21
                 f'{valor_csll};'           # 22
                 f'{valor_crf};'            # 23
                 f'{valor_inss};'           # 24
                 f'{codigo_item};'          # 25
                 f'{quantidade};'           # 26
                 f'{valor_unitario}')"""    # 27
    
    dnl = dados_nota.split(';')
    
    dados_nota_layout = (f'{dnl[0]};{dnl[1]};{dnl[2]};{dnl[3]};{dnl[4]};{dnl[5]};{dnl[6]};{dnl[7]};{dnl[8]};{dnl[9]};{dnl[10]};{dnl[11]};{dnl[12]};{dnl[13]};{dnl[14]};{dnl[15]};'
                         f'{dnl[16]};{dnl[17]};{dnl[18]};{dnl[19]};{dnl[20]};{dnl[21]};{dnl[22]};{dnl[23]};{dnl[24]};{dnl[25]};{dnl[26]};{dnl[27]}')
    
    f.write(pula_linha_txt + str(dados_nota_layout))
    f.close()
    

@_time_execution
def run():
    andamentos = 'Relatório de NFSe emitidas'
    periodo = ''
    todas_as_notas = confirm(text='Baixar todas as notas disponíveis das empresas?', buttons=('Sim', 'Não'))
    if todas_as_notas == 'Não':
        periodo = prompt(text='Qual o mês de emissão das notas? (Digite uma competência: 00/0000 ou um ano: 0000)', default='00/0000')
    
    # define e cria a pasta final das notas
    download_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\Relatório de NFSe Emitidas\\Execução\\NFSE"
    
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
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        cod_dominio, cnpj, nome, usuario, senha = empresa
        
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            driver, situacao = login(driver, usuario, senha)
            # se fizer o login realiza a consulta
            if situacao == 'ok':
                driver, resultado = consulta_notas(andamentos, periodo, download_folder, driver, cod_dominio, cnpj, nome)
                # se realizar a consulta anota na planilha
                if resultado != 'ok':
                    print(resultado)
                    _escreve_relatorio_csv(f'{cod_dominio};{cnpj};{nome};{resultado[2:]}', nome='Resumo')
                break
                
            driver.close()
        driver.close()
        
    _escreve_header_csv('CÓD. DOMÍNIO;CNPJ;NOME;CNPJ/CPF/NIF TOMADOR;NOME TOMADOR;UF TOMADOR;MUNICÍPIO TOMADOR;ENDEREÇO TOMADOR;NÚMERO DA NOTA;'
                        'SÉRIE DA NOTA;DATA DE EMISSÃO;SITUAÇÃO DA NOTA;ACUMULADOR;CFPS;VALOR DO SERVIÇO;VALOR DO DESCONTOS;VALOR DAS DEDUÇÕES;'
                        'VALOR CONTÁBIL;BASE DE CÁLCULO;ALÍQUOTA ISS;VALOR ISS;RETIDO;VALOR IRRF;VALOR PIS;VALOR COFINS;VALOR CSLL;VALOR CRF;'
                        'VALOR INSS;CODIGO ITEM;QUANTIDADE;VALOR UNITÁRIO', nome=andamentos)
    _escreve_header_csv('CÓD. DOMÍNIO;CNPJ;NOME;RESULTADO;COMP. NOTAS;VALOR TOTAL DAS NOTAS POR COMPETÊNCIA', nome='Resumo')
    
if __name__ == '__main__':
    run()
