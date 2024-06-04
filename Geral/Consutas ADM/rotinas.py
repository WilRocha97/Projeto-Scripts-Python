# -*- coding: utf-8 -*-
import selenium.common
import xhtml2pdf, fitz, sys, shutil, re, os, traceback, requests.exceptions, pdfkit
from requests import Session
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from anticaptchaofficial.hcaptchaproxyless import *
from pyautogui import alert
import comum

controle_rotinas = 'Log/Controle.txt'


def run_cndtni(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def abre_pagina_consulta(window_principal, driver):
        print('>>> Abrindo Conta Fiscal')
        window_principal['-Mensagens_2-'].update(f'Abrindo Conta Fiscal...')
        window_principal.refresh()
        
        try:
            # iteração para logar no usuário
            while not re.compile(r'>Conta Fiscal do ICMS e Parcelamento').search(driver.page_source):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return False, ''

                try:
                    button = driver.find_element(by=By.XPATH,
                                                 value='/html/body/div[2]/section/div/div/div/div[2]/div/ul/li/form/div[5]/div/a')
                    button.click()
                    time.sleep(3)
                except:
                    pass
                time.sleep(1)
        except:
            print('❗ Erro ao logar no usuário, tentando novamente')
            return driver, 'erro'
        
        # pega a url da página da consulta
        url_consulta = re.compile(r'<a href=\"(.+\d).+>Conta Fiscal do ICMS e Parcelamento').search(
            driver.page_source).group(1)
        
        # entra na página e aguarda carregar
        driver.get(url_consulta)
        
        while not comum.find_by_id('divcontainer', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return False, ''
            time.sleep(1)
        
        print('>>> Abrindo consulta de CNDNI')
        window_principal['-Mensagens_2-'].update(f'Abrindo consulta de CNDNI...')
        window_principal.refresh()
        # iteração para capturar a url da página final para poder inserir os infos da empresa e consultar
        while True:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return False, ''
            try:
                url_consulta_cndni = re.compile(r'href="(https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages.+)" tabindex="-1">Verificar Impedimentos eCND').search(driver.page_source).group(1)
                break
            except:
                pass
        
        # abre a página
        driver.get(url_consulta_cndni)
        
        return driver, 'ok'
    
    def consulta_cndni(driver, nome, cnpj, pasta_download, pasta_final):
        print('>>> Consultando.')
        while not comum.find_by_id('MainContent_txtDocumento', driver):
            time.sleep(1)
        
        # enquanto a tela com os resultados não abre, tenta logar na empresa, se der erro de CNPJ tenta novamente mais duas vezes, se não conseguir retorna erro de CNPJ
        # fora esse erro tenta logar no usuário mais 15 vezes no total sempre verificando se algúm alerta for exibido, caso seja retorna a mensagem do alerta
        contador = 0
        contador_cpf = 0
        while not comum.find_by_id('MainContent_lnkImprimirCertidaoBotao1', driver):
            while True:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return False, ''
                driver.find_element(by=By.ID, value='MainContent_txtDocumento').clear()
                time.sleep(1)
                driver.find_element(by=By.ID, value='MainContent_txtDocumento').send_keys(cnpj)
                time.sleep(1)
                driver.find_element(by=By.ID, value='MainContent_btnPesquisar').click()
                time.sleep(1)
                
                try:
                    # wait = WebDriverWait(driver, 5)
                    # alert = wait.until(expected_conditions.alert_is_present())
                    # Captura o alerta
                    alert = driver.switch_to.alert
                except:
                    alert = False
                
                if alert:
                    # Store the alert text in a variable
                    text = alert.text
                    print(f'Alert info: {text}')
                    # Press the OK button
                    alert.accept()
                    
                    if text != 'Pesquisa não autorizada. Cadastro não localizado.':
                        print(f'❌ {text}')
                        return driver, text
                    else:
                        print(f'❗ Possível erro ao digitar o CNPJ, tentando novamente.')
                        contador_cpf += 1
                    
                    if contador_cpf >= 3:
                        print(f'❌ {text}')
                        return driver, text
                else:
                    break
            
            time.sleep(1)
            contador += 1
            if contador > 15:
                print('❌ Erro ao consultar CNDNI, tentando novamente')
                return driver, 'erro'
        
        # tenta baicar o PDF com as infos da consulta enquanto verifica se deu algum erro no site se der erro ao salvar o documento o site está demorando para carregar, tenta mais 59 vezes, uma por segundo
        # enquanto tenta verifica algumas condições, o site sempre retorna um relatório caso o acesso esteja ok, se depois de 60 segundos não conseguir salvá-lo, tenta novamente.
        contador = 0
        while True:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return False
            if re.compile(r"Server Error in '/CertidaoNegativaDeb' Application.").search(driver.page_source):
                print('❌ Erro ao acessar o site, tentando novamente')
                return driver, 'erro'
            try:
                os.makedirs(pasta_download, exist_ok=True)
                while os.listdir(pasta_download) == []:
                    print('>>> Tentando baixar o documento...')
                    driver.execute_script("window.scrollBy(0,200)")
                    driver.find_element(by=By.ID, value='MainContent_lnkImprimirCertidaoBotao1').click()
                    time.sleep(3)
                
                resultado = renomeia_cndni(window_principal, nome, cnpj, pasta_download, pasta_final)
                if resultado:
                    break
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return False
            
            except:
                if re.compile(r"Ocorreu uma falha na geração do relatório!").search(driver.page_source):
                    driver.find_element(by=By.XPATH, value='/html/body/div[4]/div[3]/div/button').click()
                    print('❌ Erro ao emitir relatório, tentando novamente')
                    return driver, 'erro'
                
                print('>>> Consultando...')
                time.sleep(1)
                contador += 1
            
            if re.compile(r'Acesso negado').search(driver.page_source):
                print('❌ Acesso negado para essa empresa')
                return driver, 'Acesso negado para essa empresa'
            
            if contador > 60:
                print('❌ Erro ao consultar CNDNI, tentando novamente')
                return driver, 'erro'
        
        # volta para a página anterior para consultar a próxima empresa
        driver.execute_script('document.getElementById("MainContent_lnkVoltar").click()')
        return driver, resultado
    
    def renomeia_cndni(window_principal, nome, cnpj, pasta_download, pasta_final):
        debitos = 'não'
        time.sleep(1)
        # verifica se o arquivo está corrompido, se tiver retorna False
        for cndni in os.listdir(pasta_download):
            arq = os.path.join(pasta_download, cndni)
            if arq.endswith('.crdownload'):
                os.remove(arq)
                return False
            
            window_principal['-Mensagens_2-'].update(f'Salvando CNDNI...')
            window_principal.refresh()
            # iteração para tentar abrir o arquivo baixado, se não conseguir é porque o download não foi concluído
            while True:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return False
                try:
                    doc = fitz.open(arq, filetype="pdf")
                    break
                except:
                    doc = []
                    traceback_str = traceback.format_exc()
                    print(traceback_str)
                    print('>>> Aguardando download')
                    pass
                
                break
            
            # iteração para capturar os dados do arquivo baixado e colocar na planilha
            for page in doc:
                texto = page.get_text('text', flags=1 + 2 + 8)
                if re.compile(r'\nHá Pendências').search(texto):
                    debitos = 'sim'
                if re.compile(r'\nHá Débitos').search(texto):
                    debitos = 'sim'
                
                if debitos == 'sim':
                    debitos_lista = ''
                    
                    resumo = [('ICMS Declarado', r'ICMS Declarado\nNão há Débitos\n'),
                              ('ICMS Parcelamento', r'ICMS Parcelamento\nNão há Débitos\n'),
                              ('IPVA', r'IPVA\nNão há Débitos\n'),
                              ('ITCMD', r'ITCMD\nNão há Débitos\n'),
                              ('AIIM', r'AIIM\nNão há Débitos\n'),
                              ('ICMS Pendência', r'ICMS Pendência\nNão há Pendências\n')]
                    for item in resumo:
                        if not re.compile(item[1]).search(texto):
                            debitos_lista += ' - ' + item[0]
                    
                    resumo = [('GIA', r'\nGIA\n'),
                              ('GIA-EFD', r'\nGIA\/EFD\n'), ]
                    for item in resumo:
                        if re.compile(item[1]).search(texto):
                            debitos_lista += ' - ' + item[0]
                    
                    if debitos_lista != '':
                        doc.close()
                        pasta_debito = os.path.join(pasta_final, 'CNDNI com débitos')
                        os.makedirs(pasta_debito, exist_ok=True)
                        shutil.move(arq, os.path.join(pasta_debito, f'{nome[:30]} - {cnpj} - CNDNI Débitos{debitos_lista}.pdf'))
                        print('❗ Com Débitos')
                        return 'Com débitos'
            
            doc.close()
            pasta_sem_debito = os.path.join(pasta_final, 'CNDNI')
            os.makedirs(pasta_sem_debito, exist_ok=True)
            # move o arquivo com novo nome para melhor identificá-lo
            shutil.move(arq, os.path.join(pasta_sem_debito, f'{nome[:30]} - {cnpj} - CNDNI.pdf'))
        
        print('✔ Sem débitos')
        return 'Sem débitos'
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador'],
                                                         filtrar_celulas_em_branco=['CNPJ', 'Razao', 'Cidade'])
    if df_empresas.empty:
        return False, pasta_final
    
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    driver = ''
    
    pasta_download = os.path.join(pasta_final, 'Download')
    
    # configura o navegador
    options = comum.configura_navegador(window_principal, pasta=pasta_download, retorna_options=True)
    
    tempos = [datetime.now()]
    tempo_execucao = []
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        nome = nome.replace('/', '')
        print(cnpj)
        
        resultado = 'ok'
        while True:
            # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
            if usuario_anterior != usuario:
                
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
                try:
                    driver.close()
                except:
                    pass
                
                contador = 0
                while True:
                    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                    if cr == 'STOP':
                        alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                        try:
                            driver.close()
                        except:
                            pass
                        return False, pasta_final
                    try:
                        # abre uma nova sessão no site da fazenda
                        driver, sid = comum.new_session_fazenda_driver(window_principal, usuario, senha, perfil, retorna_driver=True, options=options)
                        if sid == 'erro_captcha':
                            return False
                        break
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
                    
                    if contador >= 5:
                        print('❌ Impossível de logar com esse usuário')
                        sid = 'Impossível de logar com esse usuário'
                        driver = 'erro'
                        break
                
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                if driver == 'erro':
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': sid},
                                                 nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                else:
                    driver, resultado = abre_pagina_consulta(window_principal, driver)
            
            # se o resultado da abertura da página de consulta for 'ok', consulta a empresa
            if resultado == 'ok':
                driver, resultado = consulta_cndni(driver, nome, cnpj, pasta_download, pasta_final)
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                if resultado != 'erro':
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': sid},
                                                 nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                driver.close()
                usuario_anterior = 'padrão'
                continue
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    driver.close()
                except:
                    pass
                return False, pasta_final
    
    try:
        driver.close()
    except:
        pass
    return True, pasta_final


def run_debitos_municipais_jundiai(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def login_jundiai(window_principal, cnpj, insc_muni, pasta_final):
        cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
        
        pasta_final_certidao = os.path.join(pasta_final, 'Certidões')
        
        # configura o navegador
        status, driver = comum.configura_navegador(window_principal, pasta=pasta_final_certidao)
        window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
        window_principal.refresh()
        
        # entra na página inicial da consulta
        url_inicio = 'https://web.jundiai.sp.gov.br/PMJ/SW/certidaonegativamobiliario.aspx'
        driver.get(url_inicio)
        time.sleep(1)
        
        # pega o sitekey para quebrar o captcha
        data = {'url': url_inicio, 'sitekey': '6LfK0igTAAAAAOeUqc7uHQpW4XS3EqxwOUCHaSSi'}
        response = comum.solve_recaptcha(data)
        if not response:
            return driver, 'recomeçar', '❌ Erro no captcha'
        
        # insere a inscrição estadual e o cnpj da empresa
        comum.send_input('DadoContribuinteMobiliario1_txtCfm', insc_muni, driver)
        comum.send_input('DadoContribuinteMobiliario1_txtNrCicCgc', cnpj, driver)
        
        # pega a id do campo do recaptcha
        id_response = re.compile(r'class=\"recaptcha\".+<textarea id=\"(.+)\" name=')
        id_response = id_response.search(driver.page_source).group(1)
        
        text_response = ''
        while not text_response == response:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar', ''
            
            print('>>> Tentando preencher o campo do captcha')
            # insere a solução do captcha via javascript
            driver.execute_script('document.getElementById("' + id_response + '").innerText="' + response + '"')
            time.sleep(2)
            try:
                text_response = re.compile(r'display: none;\">(.+)</textarea></div>')
                text_response = text_response.search(driver.page_source).group(1)
            except:
                pass
        
        # clica em consultar
        window_principal['-Mensagens_2-'].update('Consultando empresa...')
        window_principal.refresh()
        print('>>> Consultando empresa')
        driver.find_element(by=By.ID, value='btnEnviar').click()
        time.sleep(1)
        
        while not comum.find_by_id('lblContribuinte', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar', ''
            
            if comum.find_by_id('AjaxAlertMod1_lblAjaxAlertMensagem', driver):
                situacao = re.compile(r'AjaxAlertMod1_lblAjaxAlertMensagem\">(.+)</span>')
                while True:
                    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                    if cr == 'STOP':
                        return driver, 'encerrar', ''
                    
                    try:
                        situacao = situacao.search(driver.page_source).group(1)
                        break
                    except:
                        pass
                
                if situacao == 'Consta(m) pendência(s) para a emissão de certidão por meio da Internet. Dirija-se à Av. União dos Ferroviários, 1760 - ' \
                               'Centro - Jundiaí de segunda a sexta-feiras das 9h:00 às 18h:00 e aos sábados das 9h:00 às 13h:00.' \
                        or situacao == 'Consta(m) pendência(s) para emissão de certidão por meio da internet. Dirija-se à Av. União dos Ferroviários, 1760 - ' \
                                       'Centro - Jundiaí de segunda a sexta-feiras das 9h:00 às 17h:00 e aos sábados das 9h:00 às 13h:00.':
                    situacao_print = f'❗ {situacao}'
                    return driver, 'ok', situacao_print
                if situacao == 'Confirme que você não é um robô':
                    return driver, 'recomeçar', '❌ Erro ao logar na empresa'
                if situacao == 'EIX000: (-107) ISAM error:  record is locked.':
                    return driver, 'recomeçar', '❌ Erro no site'
            time.sleep(1)
        
        situacao = re.compile(r'<span id=\"lblSituacao\">(.+)</span>')
        situacao = situacao.search(driver.page_source).group(1)
        time.sleep(1)
        
        if situacao == 'Não constam débitos para o contribuinte':
            os.makedirs(pasta_final_certidao, exist_ok=True)
            
            comum.send_input('txtSolicitante', 'Escritório', driver)
            comum.send_input('txtMotivo', 'Consulta', driver)
            time.sleep(1)
            driver.find_element(by=By.ID, value='btnImprimir').click()
            
            while not os.path.exists(os.path.join(pasta_final_certidao, 'relatorio.pdf')):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver, ''
                time.sleep(1)
            while os.path.exists(os.path.join(pasta_final_certidao, 'relatorio.pdf')):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver, ''
                try:
                    arquivo = f'{cnpj_limpo} - Certidão Negativa de Débitos Municipais Jundiaí.pdf'
                    shutil.move(os.path.join(pasta_final_certidao, 'relatorio.pdf'), os.path.join(pasta_final_certidao, arquivo))
                    time.sleep(2)
                except:
                    pass
        if situacao == 'Consta(m) pendência(s) para a emissão de certidão por meio da Internet. Dirija-se à Av. União dos Ferroviários, 1760 - ' \
                       'Centro - Jundiaí de segunda a sexta-feiras das 9h:00 às 18h:00 e aos sábados das 9h:00 às 13h:00.' \
                or situacao == 'Consta(m) pendência(s) para emissão de certidão por meio da internet. Dirija-se à Av. União dos Ferroviários, 1760 - ' \
                               'Centro - Jundiaí de segunda a sexta-feiras das 9h:00 às 17h:00 e aos sábados das 9h:00 às 13h:00.':
            situacao_print = f'❗ {situacao}'
            return driver, situacao_print
        
        situacao_print = f'✔ {situacao}'
        return driver, situacao_print
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Jundiaí'])
    if df_empresas.empty:
        return False, pasta_final
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        contador = 0
        # iteração para logar na empresa e consultar, tenta 5 vezes
        while True:
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            driver, situacao, situacao_print = login_jundiai(window_principal, cnpj, insc_muni, pasta_final)
            print(situacao_print)
            
            if situacao == 'recomeçar':
                try:
                    driver.close()
                except:
                    pass
                contador += 1
                if contador >= 5:
                    break
                continue
            
            else:
                break
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            return False, pasta_final
        
        comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': situacao_print[2:]},
                                     nome=andamentos, local=pasta_final)
    
    return True, pasta_final


def run_debitos_municipais_valinhos(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def reset_table(table):
        table.attrs = None
        for tag in table.findAll(True):
            tag.attrs = None
            try:
                tag.string = tag.string.strip()
            
            except:
                pass
        
        return table
    
    def format_data(content):
        # tabelas = []
        soup = BeautifulSoup(content, 'html.parser')
        tabela_header = soup.find('table', attrs={'id': 'table6'})
        tabela_debitos = soup.find('div', attrs={'id': 'grid'}).find('table')
        tabela_totais = soup.find('table', attrs={'id': 'tableTotais'})
        
        linhas = tabela_header.findAll('td', attrs={'class': 'linhaInformacao'})
        paragrafo = ', '.join(linha.text for linha in linhas)
        
        if len(tabela_debitos.findAll('tr')) <= 1:
            return False
        # tabela_debitos.attrs = None
        for tag in tabela_debitos.findAll(True):
            # tag.attrs = None
            try:
                tag.string = tag.string.strip()
            except:
                pass
            if 'input' == tag.name:
                tag.parent.extract()
        
        tabela_totais.find('tr', attrs={'id': 'tr4'}).extract()
        tabela_totais.find('td', attrs={'id': 'td2'}).extract()
        tabela_totais.find('td', attrs={'id': 'td4'}).extract()
        # tabela_totais = reset_table(tabela_totais)
        
        with open('Assets\style.css', 'r') as f:
            css = f.read()
        
        style = "<style type='text/css'>" + css + "</style>"
        new_soup = BeautifulSoup(f'<html><head><meta charset="UTF-8">{style}</head><body><p></p></body></html>', 'html.parser')
        new_soup.body.p.append(paragrafo)
        new_soup.body.append(tabela_debitos)
        new_soup.body.append(tabela_totais)
        
        return str(new_soup)
    
    def login_valinhos(driver, cnpj, insc_muni):
        base = 'http://179.108.81.10:9081/tbw'
        url_inicio = f'{base}/loginWeb.jsp?execobj=ServicoPesquisaISSQN'
        
        driver.get(url_inicio)
        while not comum.find_by_id('span7Menu', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        button = driver.find_element(by=By.ID, value='span7Menu')
        button.click()
        
        while not comum.find_by_id('captchaimg', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        
        time.sleep(1)
        captcha = comum.solve_text_captcha(driver, 'captchaimg')
        if not captcha:
            print('Erro Login - não encontrou captcha')
            return driver, 'recomeçar'
        
        while not comum.find_by_id('input1', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input1')
        element.send_keys(insc_muni)
        
        while not comum.find_by_id('input4', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input4')
        element.send_keys(cnpj)
        
        while not comum.find_by_id('captchafield', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='captchafield')
        element.send_keys(captcha)
        
        while not comum.find_by_id('imagebutton1', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            time.sleep(1)
        button = driver.find_element(by=By.ID, value='imagebutton1')
        button.click()
        
        time.sleep(3)
        
        timer = 0
        while not comum.find_by_id('td30', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'encerrar'
            print('>>> Aguardando site')
            driver.save_screenshot(r'Log\debug_screen.png')
            time.sleep(1)
            
            if comum.find_by_id('tdMsg', driver):
                try:
                    erro = re.compile(r'id=\"tdMsg\".+tipo=\"td\">(.+)\.').search(driver.page_source).group(1)
                    print(f'❌ {erro}')
                    return driver, erro
                except:
                    pass
            timer += 1
            if timer >= 10:
                print(f'❌ Erro no login')
                return driver, 'recomeçar'
        
        return driver, 'ok'
    
    def consulta_valinhos(driver, cnpj, insc_muni, nome, andamentos, pasta_final):
        print('>>> Consultando empresa')
        html = driver.page_source.encode('utf-8')
        try:
            driver.save_screenshot(r'Log\debug_screen.png')
            str_html = format_data(html)
            
            if str_html:
                print('>>> Gerando arquivo')
                os.makedirs(os.path.join(pasta_final, 'Arquivos'), exist_ok=True)
                nome_arq = ';'.join([cnpj, 'INF_FISC_REAL', 'Debitos Municipais'])
                with open(os.path.join(pasta_final, 'Arquivos', nome_arq + '.pdf'), 'w+b') as pdf:
                    xhtml2pdf.pisa.showLogging()
                    xhtml2pdf.pisa.CreatePDF(str_html, pdf)
                    print('❗ Arquivo gerado')
                comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': 'Com débitos'},
                                             nome=andamentos, local=pasta_final)
            else:
                comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': 'Sem débitos'},
                                             nome=andamentos, local=pasta_final)
                print('✔ Não há débitos')
        except:
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': 'Erro na geração do PDF'},
                                         nome=andamentos, local=pasta_final)
            driver.save_screenshot(r'Log\debug_screen.png')
            print('❌ Erro na geração do PDF')
        
        return driver, True
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Valinhos'])
    if df_empresas.empty:
        return False, pasta_final
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        try:
            insc_muni = int(insc_muni)
        except:
            continue
        
        while True:
            status, driver = comum.configura_navegador(window_principal)
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            driver, resultado = login_valinhos(driver, cnpj, insc_muni)
            
            if resultado == 'encerrar':
                break
            if resultado != 'recomeçar':
                if resultado == 'ok':
                    driver, resultado = consulta_valinhos(driver, cnpj, insc_muni, nome, andamentos, pasta_final)
                else:
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': resultado},
                                                 nome=andamentos, local=pasta_final)
                break
            
            window_principal['-Mensagens_2-'].update('Erro ao logar na empresa, tentando novamente...')
            window_principal.refresh()
        
        try:
            driver.close()
        except:
            pass
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            return False, pasta_final
    
    return True, pasta_final


def run_debitos_municipais_vinhedo(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final):
        # configura o navegador
        status, driver = comum.configura_navegador(window_principal, pasta=pasta_final)
        
        url_inicio = 'http://vinhedomun.presconinformatica.com.br/certidaoNegativa.jsf?faces-redirect=true'
        driver.get(url_inicio)
        
        contador = 1
        while not comum.find_by_id(f'homeForm:panelCaptcha:j_idt{str(contador)}', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return False, 'encerrar'
            contador += 1
            time.sleep(0.2)
        
        # resolve o captcha
        captcha = comum.solve_text_captcha(driver, f'homeForm:panelCaptcha:j_idt{str(contador)}')
        if not captcha:
            try:
                driver.close()
            except:
                pass
            return False, 'recomeçar'
        
        # espera o campo do tipo da pesquisa
        while not comum.find_by_id('homeForm:inputTipoInscricao_label', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return False, 'encerrar'
            time.sleep(1)
        # clica no menu
        driver.find_element(by=By.ID, value='homeForm:inputTipoInscricao_label').click()
        
        # espera o menu abrir
        while not comum.find_by_path('/html/body/div[6]/div/ul/li[2]', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return False, 'encerrar'
            time.sleep(1)
        # clica na opção "Mobiliário"
        driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/ul/li[2]').click()
        
        if not captcha:
            print('Erro Login - não encontrou captcha')
            try:
                driver.close()
            except:
                pass
            return False, 'recomeçar'
        
        time.sleep(1)
        # clica na barra de inscrição e insere
        driver.find_element(by=By.ID, value='homeForm:inputInscricao').click()
        time.sleep(2)
        driver.find_element(by=By.ID, value='homeForm:inputInscricao').send_keys(insc_muni)
        comum.send_input(f'homeForm:panelCaptcha:j_idt{str(contador + 3)}', captcha, driver)
        time.sleep(2)
        
        # clica no botão de pesquisar
        driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[2]/div[5]/a/div').click()
        
        print('>>> Consultando')
        window_principal['-Mensagens_2-'].update('Consultando empresa...')
        window_principal.refresh()
        while 'dados-contribuinte-inscricao">0000000000' not in driver.page_source:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return False, 'encerrar'
            
            if comum.find_by_path('/html/body/div[1]/div[5]/div[2]/div[1]/div/ul/li/span', driver):
                padraozinho = re.compile(r'confirmationMessage\" class=\"ui-messages ui-widget\" aria-live=\"polite\"><div '
                                         r'class=\"ui-messages-error ui-corner-all\"><span class=\"ui-messages-error-icon\"></span><ul><li><span '
                                         r'class=\"ui-messages-error-summary\">(.+).</span></li></ul></div></div><div class=\"right\"><button id=\"j_idt')
                situacao = padraozinho.search(driver.page_source).group(1)
                if situacao == 'Letras de segurança inválidas' or situacao == 'Por favor, informe o(a) Inscrição Cadastral':
                    driver.close()
                    return False, 'recomeçar'
                
                situacao_print = f'❌ {situacao}'
                # print(driver.page_source)
                driver.close()
                return situacao, situacao_print
            
            time.sleep(1)
        
        situacao = salvar_guia_vinhedo(window_principal, driver, cnpj, pasta_final)
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            try:
                driver.close()
            except:
                pass
            return False, 'encerrar'
        
        situacao_print = f'✔ {situacao}'
        return situacao, situacao_print
    
    def salvar_guia_vinhedo(window_principal, driver, cnpj, pasta_final):
        window_principal['-Mensagens_2-'].update('Buscando Certidão Negativa...')
        window_principal.refresh()
        time.sleep(1)
        while True:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return
            try:
                print('>>> Localizando Certidão')
                driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[4]/div[1]/a[1]/div').click()
                break
            except:
                time.sleep(1)
                pass
        
        while '<object' not in driver.page_source:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        
        url_pdf = re.compile(r'(/impressao\.pdf.+)\" height=\"500px\" width=\"100%\">').search(driver.page_source).group(1)
        
        window_principal['-Mensagens_2-'].update('Salvando Certidão Negativa...')
        window_principal.refresh()
        print('>>> Salvando Certidão Negativa')
        driver.get('https://vinhedomun.presconinformatica.com.br/' + url_pdf)
        time.sleep(1)
        while True:
            try:
                shutil.move(os.path.join(pasta_final, 'impressao.pdf'), os.path.join(pasta_final, f'{cnpj} Certidão Negativa de Débitos Municipais Vinhedo.pdf'))
                break
            except:
                pass
        driver.close()
        
        return 'Certidão negativa salva'
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Vinhedo'])
    if df_empresas.empty:
        return False, pasta_final
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        while True:
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            
            situacao, situacao_print = pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final)
            if situacao_print == 'encerrar':
                break
                
            if situacao_print != 'recomeçar':
                print(situacao_print)
                if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
                    alert(f'❗ Rotina "{andamentos}":\n\nDesculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.')
                    return False, pasta_final
                
                comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': situacao},
                                             nome=andamentos, local=pasta_final)
                break
            
            window_principal['-Mensagens_2-'].update('Erro ao logar na empresa, tentando novamente...')
            window_principal.refresh()
            
            
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            return False, pasta_final
        
    return True, pasta_final


def run_debitos_estaduais(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    situacoes_debitos_estaduais = {
        'C': '✔ Certidão sem pendencias',
        'S': '❗ Nao apresentou STDA',
        'G': '❗ Nao apresentou GIA',
        'E': '❌ Nao baixou arquivo',
        'T': '❗ Transporte de Saldo Credor Incorreto',
        'P': '❗ Pendencias',
        'I': '❗ Pendencias GIA',
    }
    
    def confere_pendencias(pagina):
        print('>>> Verificando pendencias')
        
        id_base = 'MainContent_'
        soup = BeautifulSoup(pagina.content, 'html.parser')
        pendencia = [
            # soup.find('span', attrs={'id':f'{id_base}lblMsgErroParcelamento'}).text != '', # parce
            soup.find('span', attrs={'id': f'{id_base}lblMsgErroResultado'}).text != '',  # deb inscritos
            soup.find('span', attrs={'id': f'{id_base}lblMsgErroFiltro'}).text != '',  # ocorrências
        ]
        
        if all(pendencia):
            return situacoes_debitos_estaduais['C']
        
        situacao = []
        if not pendencia[0]:
            attrs = {'id': f'{id_base}rptListaDebito_lblValorTotalDevido'}
            deb = soup.find('span', attrs=attrs)
            deb = 'R$ 0,00' if not deb else deb.text
            
            pendencia[0] = float(deb[3:].replace('.', '').replace(',', '.')) == 0
            if all(pendencia):
                return situacoes_debitos_estaduais['C']
            
            if re.findall(r'GIA-1/1', str(soup)):
                situacao.append(situacoes_debitos_estaduais['I'])
            elif re.findall(r'GIA ST-1/1', str(soup)):
                situacao.append(situacoes_debitos_estaduais['I'])
            elif re.findall(r'MainContent_rptListaDebito_rptDetalheDebito_0_lblValorOrigem_0\">GIA<', str(soup)):
                situacao.append(situacoes_debitos_estaduais['I'])
            else:
                situacao.append(situacoes_debitos_estaduais['P'])
        
        if not pendencia[1]:
            tabela = soup.find('table', attrs={'id': f'{id_base}gdvResultado'})
            if not tabela:
                return '❌ Erro ao analisar GIA/STDA'
            
            linhas = tabela.find_all('tr')
            if not linhas:
                return '❌ Erro ao analisar GIA/STDA'
            
            for linha in linhas:
                if 'Não apresentou GIA' in linha.text:
                    situacao.append(situacoes_debitos_estaduais['G'])
                    break
                
                if 'Não apresentou STDA' in linha.text:
                    situacao.append(situacoes_debitos_estaduais['S'])
                    break
                
                if 'Transporte de Saldo Credor Incorreto' in linha.text:
                    situacao.append(situacoes_debitos_estaduais['T'])
                    break
        
        return ' e '.join(situacao)
    
    def consulta_deb_estaduais(window_principal, pasta_final, cnpj, cidade, s, s_id):
        if cidade == 'Jundiaí':
            pasta_final = os.path.join(pasta_final, 'Relatórios Filial Jundiaí')
        else:
            pasta_final = os.path.join(pasta_final, 'Relatórios Matriz Valinhos')
        
        window_principal['-Mensagens_2-'].update(f'Consultando débitos...')
        window_principal.refresh()
        print('>>> Consultando débitos')
        url_base = 'https://www10.fazenda.sp.gov.br/ContaFiscalIcms/Pages'
        url_pesquisa = f'{url_base}/SituacaoContribuinte.aspx?SID={s_id}'
        
        # formata o cnpj colocando os separadores
        f_cnpj = f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
        
        erro = True
        state = ''
        validation = ''
        generator = ''
        while erro:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return False
            try:
                res = s.get(url_pesquisa)
                time.sleep(2)
                
                state, generator, validation = comum.get_info_post(res.content)
                erro = False
            except:
                erro = True
        
        info = {
            '__EVENTTARGET': 'ctl00$MainContent$btnConsultar',
            '__EVENTARGUMENT': '', '__VIEWSTATE': state,
            '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION': validation,
            'ctl00$MainContent$hdfCriterioAtual': '',
            'ctl00$MainContent$ddlContribuinte': 'CNPJ',
            'ctl00$MainContent$txtCriterioConsulta': f_cnpj
        }
        
        erro = True
        res = ''
        soup = ''
        while erro:
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return False
            
            try:
                res = s.post(url_pesquisa, info)
                soup = BeautifulSoup(res.content, 'html.parser')
                
                state, generator, validation = comum.get_info_post(res.content)
                erro = False
            except:
                print('Erro ao gerar info para a consulta')
                
                erro = True
        
        info['__EVENTTARGET'] = 'ctl00$MainContent$lkbImpressao'
        info['__VIEWSTATE'] = state
        info['__EVENTVALIDATION'] = validation
        info['__VIEWSTATEGENERATOR'] = generator
        info['ctl00$MainContent$hdfCriterioAtual'] = f_cnpj
        info['ctl00$MainContent$txtCriterioConsulta'] = f_cnpj
        
        id_base = 'MainContent_'
        attrs = {'id': f'{id_base}lblMensagemDeErro'}
        check = soup.find('span', attrs=attrs)
        if check:
            return '❗ ' + check.text.strip()
        
        try:
            situacao = confere_pendencias(res)
            if situacao == situacoes_debitos_estaduais['C']:
                return situacao
        except AttributeError:
            return "Nao identificada"
        
        impressao = s.post(url_pesquisa, info)
        if impressao.headers.get('content-disposition', ''):
            nome = f"{cnpj};INF_FISC_REAL;Debitos Estaduais - {situacao.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')}.pdf"
            comum.download_file(nome, impressao, pasta=pasta_final)
        else:
            situacao = situacoes_debitos_estaduais['E']
        
        return situacao
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador'],
                                                         filtrar_celulas_em_branco=['CNPJ', 'Razao', 'Cidade'])
    
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    s = False
    
    if df_empresas.empty:
        return False, pasta_final
    
    tempos = [datetime.now()]
    tempo_execucao = []
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        cnpj = str(cnpj)
        print(cnpj)
        
        # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
        if usuario_anterior != usuario:
            # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
            if s: s.close()
            
            # abre uma nova sessão
            s = Session()
            
            contador = 0
            # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
            while True:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    if s: s.close()
                    return False, pasta_final
                
                if contador >= 3:
                    cookies = 'erro'
                    sid = 'Erro ao logar na empresa'
                    break
                try:
                    cookies, sid = comum.new_session_fazenda_driver(window_principal, usuario, senha, perfil)
                    break
                except:
                    print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                if s: s.close()
                return False, pasta_final
            
            time.sleep(1)
            # se não salvar os cookies vai para o próximo dado
            if cookies == 'erro' or not cookies:
                texto = f'{cnpj};{sid}'
                usuario_anterior = 'padrão'
                
                comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': texto},
                                             nome=andamentos, local=pasta_final)
                continue
            
            # adiciona os cookies do login da sessão por request no web driver
            for cookie in cookies:
                s.cookies.set(cookie['name'], cookie['value'])
        
        # se não retornar a id da sessão do web driver fecha a sessão por request
        if not sid:
            situacao = '❌ Erro no login'
            usuario_anterior = 'padrão'
            s.close()
        
        # se retornar a id da sessão do web driver executa a consulta
        else:
            # retorna o resultado da consulta
            situacao = consulta_deb_estaduais(window_principal, pasta_final, cnpj, cidade, s, sid)
            # guarda o usuario da execução atual
            usuario_anterior = usuario
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                if s: s.close()
                return False, pasta_final
        
        # escreve na planilha de andamentos o resultado da execução atual
        texto = f'{cnpj};{str(situacao[2:])}'
        try:
            texto = texto.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': texto},
                                        nome=andamentos, local=pasta_final)
        except:
            raise Exception(f"Erro ao escrever esse texto: {texto}")
        print(situacao)
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            if s: s.close()
            return False, pasta_final
    
    if s: s.close()
    return True, pasta_final


def run_debitos_divida_ativa(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def verifica_debitos(pagina):
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # debito = soup.find('span', attrs={'id': 'consultaDebitoForm:dataTable:tb'})
        soup = soup.prettify()
        # print(soup)
        try:
            request_verification = re.compile(r'(consultaDebitoForm:dataTable:tb)').search(soup).group(1)
            return request_verification
        except:
            try:
                re.compile(r'(Nenhum resultado com os critérios de consulta)').search(soup).group(1)
                return False
            except:
                print(soup)
                raise Exception('que merda em')
                
    def limpa_registros(html):
        # pega todas as tags <tr>
        linhas = html.findAll('tr')
        for i in linhas:
            # pega as tags <td> dentro das tags <tr>
            colunas = i.findAll('td')
            if len(colunas) >= 7:
                # Remove informações de pagamentos
                if colunas[5].text.startswith('Liquidar'):
                    for a in colunas[5].findAll('a'):
                        a.extract()
                # Remove informações de emissão de GARE
                if colunas[6].find('table'):
                    for t in colunas[6].findAll('table'):
                        t.extract()
        
        return html
    
    def salva_pagina(pagina, cnpj, empresa, andamentos, pasta_final, compl=''):
        pasta_documentos = os.path.join(pasta_final, 'Documentos')
        os.makedirs(pasta_documentos, exist_ok=True)
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # captura o formulário com os dados de débitos
        formulario = soup.find('form', attrs={'id': 'consultaDebitoForm'})
        
        # pega todas as imágens e deleta do código
        imagens = formulario.findAll('img')
        for img in imagens:
            img.extract()
        
        # pega todos os botões e deleta do código
        botoes = formulario.findAll('input', attrs={'type': 'submit'})
        for botao in botoes:
            botao.extract()
        
        # abre o arquivo .css já criado com o layout parecido com o do formulário do site original
        with open('Assets\\style.css', 'r') as f:
            css = f.read()
        
        # insere o css criado no código do site,
        # pois quando é realizado uma requisição ela só pega o código html
        # isso é feito para gerar um PDF a partir do código html
        # e ele precisa estar estilizado para não prejudicar a leitura das informações
        style = "<style type='text/css'>" + css + "</style>"
        html = BeautifulSoup(f'<html><head><meta charset="UTF-8">{style}</head><body></body></html>', 'html.parser')
        
        # pega cada parte do formulário se existir
        # cabeçalho
        try:
            parte = formulario.find('div', attrs={'id': 'consultaDebitoForm:j_id127_header'})
            html.body.append(parte)
        except:
            pass
        # devedor
        try:
            parte = formulario.find('span', attrs={'id': 'consultaDebitoForm:consultaDevedor'}).find('table')
            html.body.append(parte)
        except:
            pass
        # titulo tabela
        try:
            parte = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTableDebitosRelativo'})
            html.body.append(parte)
        except:
            pass
        # tabela principal
        try:
            new_tag = BeautifulSoup('<table class="main-table"></table>', 'html.parser')
            head = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('thead')
            body = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('tbody')
            body = limpa_registros(body)
            footer = formulario.find('table', attrs={'id': 'consultaDebitoForm:dataTable2'}).find('tfoot')
            new_tag.table.append(head)
            new_tag.table.append(body)
            new_tag.table.append(footer)
            html.body.append(new_tag)
        except:
            pass
        
        # configura o nome do PDF
        debito = f' - {compl}' if compl else ''
        nome = os.path.join(pasta_documentos, f'{cnpj};INF_FISC_REAL;Procuradoria Federal{debito.replace(" / ", " ").replace("/", " ")}.pdf')
        # cria o PDF a partir do HTML usando a função "pisa"
        with open(nome, 'w+b') as pdf:
            # pisa.showLogging()
            xhtml2pdf.pisa.CreatePDF(str(html), pdf)
        print('✔ Arquivo salvo')
        comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': empresa, 'RESULTADO': f'Empresa com débitos', 'COMPLEMENTO': compl},
                                    nome=andamentos, local=pasta_final)
        
        return True
    
    def inicia_sessao():
        # configura navegador
        status, driver = comum.configura_navegador(window_principal)
        
        url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
        driver.get(url)
        
        # gera o token para passar pelo captcha
        recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
        token = comum.solve_recaptcha(recaptcha_data)
        if not token:
            return False, 'erro_captcha'
        
        # clica para mudar a opção de pesquisa para CNPJ
        driver.find_element(by=By.ID, value='consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa').click()
        driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div/div[2]/div/div[3]/div/div[2]/div[2]/span/form/span/div/div[2]/table/tbody/tr/td[1]/div/div/span/select/option[2]').click()
        time.sleep(1)
        
        # insere o CNPJ
        driver.find_element(by=By.ID, value='consultaDebitoForm:decTxtTipoConsulta:cnpj').send_keys('00656064000145')
        # executa um comando JS para inserir o token do captcha no campo oculto
        driver.execute_script("document.getElementById('g-recaptcha-response').innerText='" + token + "'")
        time.sleep(1)
        
        nome_botao_consulta = re.compile(r'type=\"submit\" name=\"(consultaDebitoForm:j_id\d+)\" value=\"Consultar\"').search(driver.page_source).group(1)
        # executa um comando JS para clicar no botão de consultar
        driver.execute_script("document.getElementsByName('" + nome_botao_consulta + "')[0].click()")
        time.sleep(1)
        return driver, nome_botao_consulta
    
    def consulta_debito(window_principal, s, nome_botao_consulta, empresa, andamentos, pasta_final):
        window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
        window_principal.refresh()
        
        cnpj, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        cnpj_formatado = re.sub(r'^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$', r'\1.\2.\3/\4-\5', cnpj)
        
        print(cnpj_formatado)
        
        url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
        # str_cnpj = f"{cnpj[0:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
        pagina = s.get(url)
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # tenta capturar o viewstate do site, pois ele é necessário para a requisição e é um dado que muda a cada acesso
        try:
            viewstate = soup.find(id="javax.faces.ViewState").get('value')
        except Exception as e:
            print('❌ Não encontrou viewState')
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj_formatado, 'NOME': nome, 'RESULTADO': e},
                                         nome=andamentos, local=pasta_final)
            print(e)
            print(soup)
            s.close()
            return False

        # gera o token para passar pelo captcha
        recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
        token = comum.solve_recaptcha(recaptcha_data)
        if not token:
            return 'erro_captcha'
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            return
        
        soup = soup.prettify()
        j_id_1 = re.compile(r'consultaDebitoForm:decLblTipoConsulta:j_id(\d+)').search(soup).group(1)
        
        # Troca opção de pesquisa para CNPJ
        info = {
            'AJAXREQUEST': '_viewRoot',
            'consultaDebitoForm': 'consultaDebitoForm',
            'consultaDebitoForm:decTxtTipoConsulta:cdaEtiqueta': '',
            'g-recaptcha-response': '',
            'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
            'consultaDebitoForm:modalDadosCartorioOpenedState': '',
            'javax.faces.ViewState': viewstate,
            'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
            'ajaxSingle': 'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa',
            f'consultaDebitoForm:decLblTipoConsulta:j_id{j_id_1}': f'consultaDebitoForm:decLblTipoConsulta:j_id{j_id_1}'
        }

        s.post(url, info)
        
        # Consulta o cnpj
        info = {
            'consultaDebitoForm': 'consultaDebitoForm',
            'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
            'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj_formatado,
            'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
            'g-recaptcha-response': token,
            nome_botao_consulta: 'Consultar',
            'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
            'consultaDebitoForm:modalDadosCartorioOpenedState': '',
            'javax.faces.ViewState': viewstate
        }
        pagina = s.post(url, info)
        window_principal['-Mensagens_2-'].update('Verificando débitos...')
        window_principal.refresh()
        debitos = verifica_debitos(pagina)
        
        # se não tiver débitos, anota na planilha e retorna
        if not debitos:
            print('✔ Sem débitos')
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj_formatado, 'NOME': nome, 'RESULTADO': 'Empresa sem debitos'},
                                         nome=andamentos, local=pasta_final)
        # se tiver débitos pega quantos débitos tem para montar a PDF
        else:
            print('❗ Com débitos')
            soup = BeautifulSoup(pagina.content, 'html.parser')
            tabela = soup.find('tbody', attrs={'id': 'consultaDebitoForm:dataTable:tb'})
            linhas = tabela.find_all('a')
            # captura o viewstate do site, pois ele é necessário para a requisição e é um dado que muda a cada acesso
            viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
            
            # para cada linha da tabela de débitos cria um PDF e após criar o PDF de cada débito, retorna
            for index, linha in enumerate(linhas):
                tipo = linha.get('id')
                info = {
                    'consultaDebitoForm': 'consultaDebitoForm',
                    'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
                    'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
                    'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
                    'g-recaptcha-response': token,
                    'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
                    'consultaDebitoForm:modalDadosCartorioOpenedState': '',
                    'javax.faces.ViewState': viewstate,
                    f'consultaDebitoForm:dataTable:{index}:lnkConsultaDebito': tipo
                }
                pagina = s.post(url, info)
                # cria o PDF do débito
                window_principal['-Mensagens_2-'].update(f'Débitos encontrados, gerando documentos ({index + 1} de {len(linhas)})...')
                window_principal.refresh()
                
                salva_pagina(pagina, cnpj, empresa, andamentos, pasta_final, linha.text)
                
                # Retorna para tela de consulta para gerar o PDF do outro débito
                viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
                info = {
                    'consultaDebitoForm': 'consultaDebitoForm',
                    'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
                    'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
                    'consultaDebitoForm:decTxtTipoConsulta:tiposDebitosCnpj': 0,
                    'g-recaptcha-response': token,
                    'consultaDebitoForm:j_id260': 'Voltar',
                    'consultaDebitoForm:modalSelecionarDebitoOpenedState': '',
                    'consultaDebitoForm:modalDadosCartorioOpenedState': '',
                    'javax.faces.ViewState': viewstate
                }
                s.post(url, info)
                viewstate = soup.find('input', attrs={'id': "javax.faces.ViewState"}).get('value')
        
        return s
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade'], colunas_filtro=False, palavras_filtro=False)
    if df_empresas.empty:
        return False, pasta_final
    
    # inicia uma sessão com webdriver para gerar cookies no site e garantir que as requisições funcionem depois
    driver, nome_botao_consulta = inicia_sessao()
    if nome_botao_consulta == 'erro_captcha':
        return False, pasta_final

    # armazena os cookies gerados pelo webdriver
    cookies = driver.get_cookies()
    driver.quit()
    
    # inicia uma sessão para as requisições
    with Session() as s:
        # pega a lista de cookies armazenados e os configura na sessão
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])
        
        for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
            # printa o índice da empresa que está sendo executada
            tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
            
            # enquanto a consulta não conseguir ser realizada e tenta de novo
            while True:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    break
                    
                try:
                    s = consulta_debito(window_principal, s, nome_botao_consulta, empresa, andamentos, pasta_final)
                    if s == 'erro_captcha':
                        return False, pasta_final
                    break
                except requests.exceptions.ConnectionError:
                    pass
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                return False, pasta_final
    
    return True, pasta_final


def run_debitos_divida_ativa_uniao(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    ultima_divida = 'Log/ultima_divida.txt'
    
    def click(driver, divida):
        # função para clicar em elementos via ID
        print('>>> Baixando arquivo')
        contador_2 = 0
        while True:
            try:
                driver.find_element(by=By.ID, value=divida).click()
                break
            except:
                contador_2 += 1
            
            if contador_2 > 5:
                return False
        
        return True
    
    def download_divida(driver, divida, descricao, tipo, processo, inscricao, andamentos, pasta_final_download, pasta_final_dividas):
        # formata o texto da descrição da dívida e define as pastas de download do arquivo e a pasta final do PDF
        descricao = descricao.replace('/', ' ')
        
        # cria as pastas de download e a pasta final do arquivo
        os.makedirs(pasta_final_download, exist_ok=True)
        
        # limpa a pasta de download para não ter conflito com novos arquivos
        for arquivo in os.listdir(pasta_final_download):
            os.remove(os.path.join(pasta_final_download, arquivo))
        
        # tenta clicar em download até conseguir
        while not click(driver, divida):
            time.sleep(5)
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver, 'Finalizado', False
        
        # se conseguir, mas não aparecer nada na pasta, continua tentando baixar até o limite de 6 tentativas
        contador_3 = 0
        contador_4 = 0
        while os.listdir(pasta_final_download) == []:
            if contador_3 > 10:
                while not click(driver, divida):
                    time.sleep(5)
                contador_4 += 1
                contador_3 = 0
            if contador_4 > 5:
                return driver, 'erro', True
            time.sleep(2)
            contador_3 += 1
        
        # verifica se os arquivos baixados estão com erro ou ainda não terminaram de baixar
        for arquivo in os.listdir(pasta_final_download):
            print(arquivo)
            # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
            while re.compile(r'.tmp').search(arquivo):
                try:
                    os.remove(os.path.join(pasta_final_download, arquivo))
                    return driver, 'ok', True
                except:
                    pass
                for arq in os.listdir(pasta_final_download):
                    arquivo = arq
            
            while re.compile(r'crdownload').search(arquivo):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver, 'Finalizado', False
                
                print('>>> Aguardando download...')
                time.sleep(3)
                for arq in os.listdir(pasta_final_download):
                    arquivo = arq
            
            # se não tiver problema com o arquivo baixado, tenta converter
            for arq in os.listdir(pasta_final_download):
                resultado = converte_html_pdf(pasta_final_download, pasta_final_dividas, arq, descricao, tipo, processo, inscricao, andamentos)
                if resultado == 'erro':
                    return driver, 'erro', False
                time.sleep(2)
            
            # limpa a pasta de download caso fique algum arquivo nela
            for arquivo in os.listdir(pasta_final_download):
                os.remove(os.path.join(pasta_final_download, arquivo))
                
        return driver, 'ok', False
    
    def converte_html_pdf(download_folder, pasta_final_dividas, arquivo, descricao, tipo, processo, inscricao, andamentos):
        # Defina o caminho para o arquivo HTML e PDF
        html_path = os.path.join(download_folder, arquivo)
        
        # Estrutura básica do HTML que queremos adicionar
        estrutura_inicial = """<!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
        </head>
        <body>
        """
        
        while True:
            try:
                # Leitura do conteúdo do arquivo existente
                with open(html_path, 'r', encoding='utf-8') as file:
                    conteudo_existente = file.read()
                break
            except PermissionError:
                pass
        
        # Adicionando a estrutura inicial ao conteúdo existente
        novo_conteudo = estrutura_inicial + conteudo_existente + "\n</body>\n</html>"
        while True:
            try:
                # Salvando o conteúdo modificado de volta ao arquivo
                with open(html_path, 'w', encoding='utf-8') as file:
                    file.write(novo_conteudo)
                break
            except PermissionError:
                pass
        
        # define o novo nome do arquivo PDF
        novo_arquivo, descricao, nome, cpf_cnpj = pega_info_arquivo(html_path, descricao)
        if not novo_arquivo:
            return 'erro'
        
        # coloca na pasta de ativas ou extintas
        if re.compile('EXTINTA').search(descricao) or re.compile('NEGOCIAD').search(descricao) or re.compile('Parcelad').search(descricao):
            pasta_final_dividas_ok = f'{pasta_final_dividas} EXTINTAS OU NEGOCIADAS'
        else:
            pasta_final_dividas_ok = f'{pasta_final_dividas} ATIVAS'
        os.makedirs(pasta_final_dividas_ok, exist_ok=True)
        
        # Defina o caminho para o arquivo HTML e PDF
        pdf_path = os.path.join(pasta_final_dividas_ok, novo_arquivo)
        
        # Localização do executável do wkhtmltopdf
        wkhtmltopdf_path = r'wkhtmltopdf\bin\wkhtmltopdf.exe'
        
        # Configuração do pdfkit para utilizar o wkhtmltopdf
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        
        options = {
            'page-size': 'Letter',
            'margin-top': '0.75in',
            'margin-right': '0.75in',
            'margin-bottom': '0.75in',
            'margin-left': '0.75in',
            'encoding': "UTF-8",
            'custom-header': [
                ('Accept-Encoding', 'gzip')
            ],
            'cookie': [
                ('cookie-empty-value', '""'),
                ('cookie-name1', 'cookie-value1'),
                ('cookie-name2', 'cookie-value2')
            ],
            'no-outline': None
        }
        
        # tenta converter o PDF, se não conseguir anota na planilha o nome do arquivo em html que deu erro na conversão
        # e coloca ele em uma pasta separada para arquivos em html
        try:
            pdfkit.from_file(html_path, pdf_path, configuration=config, options=options)
        except:
            comum.escreve_relatorio_xlsx({'CNPJ': cpf_cnpj, 'NOME': nome, 'TIPO':tipo, 'PROCESSO':processo, 'INSCRIÇÃO':inscricao, 'RESULTADO': f'Erro ao converter arquivo, arquivo HTML salvo: {novo_arquivo.replace(".pdf", ".html")}'},
                                         nome=andamentos, local=pasta_final)
            pasta_final_dividas_html = f'{pasta_final_dividas} Arquivos em HTML'
            os.makedirs(pasta_final_dividas_html, exist_ok=True)
            shutil.move(html_path, os.path.join(pasta_final_dividas_html, novo_arquivo.replace('.pdf', '.html')))
            print(f'❗ Erro ao converter arquivo\n')
            return False
        
        # anota o arquivo que acabou de baixar e converter
        print(f'✔ {novo_arquivo}\n')
        comum.escreve_relatorio_xlsx({'CNPJ': cpf_cnpj, 'NOME': nome, 'TIPO': tipo, 'PROCESSO': processo, 'INSCRIÇÃO': inscricao, 'RESULTADO': descricao},
                                     nome=andamentos, local=pasta_final)
        
        return True
    
    def pega_info_arquivo(html_path, descricao):
        # abrir o arquivo html
        while True:
            print('>>> Tentando abrir o novo arquivo html...')
            try:
                # Salvando o conteúdo modificado de volta ao arquivo
                with open(html_path, 'r', encoding='utf-8') as file:
                    html = file.read()
                break
            except PermissionError:
                pass
            
        # extrair o texto do arquivo html usando BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text()
        
        # pega as infos da empresa e a inscrição do documento porque tem empresas com mais de um arquivo
        if re.compile(r'Tempo de conexão esgotado').search(text):
            return False, 'Erro ao acessar o arquivo no portal REGULARIZE', False, False
        
        try:
            nome_inteiro = re.compile(r'RFB\n\n\n\nNome:\n(.+)').search(text).group(1)
        except:
            try:
                nome_inteiro = re.compile(r'Devedor Principal:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
            except:
                try:
                    nome_inteiro = re.compile(r'Devedor principal: (.+)/CNPJ:').search(text).group(1)
                except:
                    try:
                        nome_inteiro = re.compile(r'Devedor Principal:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        try:
            cpf_cnpj = re.compile(r'CNPJ/CPF:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
        except:
            try:
                cpf_cnpj = re.compile(r'CNPJ:.+(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)').search(text).group(1)
            except:
                try:
                    cpf_cnpj = re.compile(r'CNPJ:.+(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(text).group(1)
                except:
                    try:
                        cpf_cnpj = re.compile(r'CNPJ/CPF:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        try:
            inscricao = re.compile(r'Inscrição:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
        except:
            try:
                inscricao = re.compile(r'Nº inscrição:\s+(\d.+\d).+tuação da inscrição').search(text).group(1)
            except:
                try:
                    sem_inscricao = re.compile(r'Nº inscrição:\s+Situação da inscrição').search(text)
                    if sem_inscricao:
                        inscricao = 'Não foi possível recuperar as informações da inscrição'
                    else:
                        raise Exception
                except:
                    try:
                        inscricao = re.compile(r'N\.º Inscrição:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        
        if descricao == 'SEM DESCRIÇÃO':
            try:
                descricao = re.compile(r'Situação da inscrição:\s\s(.+)\s\sInformações gerais').search(text).group(1)
            except:
                try:
                    descricao = re.compile(r'Situação da inscrição: Benefício Fiscal - (.+)\s\sInformações gerais').search(text).group(1)
                except:
                    print(text)
                    raise Exception('Errei fui muleke')
        
        # formata os textos
        nome_inteiro = nome_inteiro.replace('&amp;', '&')
        inscricao = inscricao.replace('.', '').replace('-', '').replace(' ', '')
        nome = nome_inteiro[:20]
        nome = nome.replace('/', '').replace("-", "")
        cpf_cnpj = cpf_cnpj.replace('.', '').replace('/', '').replace('-', '')
        
        nome_pdf = f'{nome} - {cpf_cnpj} - {descricao}-{inscricao}.pdf'
        nome_pdf = nome_pdf.replace('  ', ' ').replace('/', '-')
        
        return nome_pdf, descricao, nome_inteiro, cpf_cnpj
    
    def login_sieg(driver):
        try:
            driver.get('https://auth.sieg.com/')
        
            print('>>> Acessando o site')
            time.sleep(1)
            
            dados = "V:\\Setor Robô\\Scripts Python\\_comum\\DadosVeri-SIEG.txt"
            f = open(dados, 'r', encoding='latin-1')
            user = f.read()
            user = user.split('/')
            
            # inserir o email no campo
            driver.find_element(by=By.ID, value='txtEmail').send_keys(user[0])
            time.sleep(1)
            
            # inserir a senha no campo
            driver.find_element(by=By.ID, value='txtPassword').send_keys(user[1])
            time.sleep(1)
            
            # clica em acessar
            driver.find_element(by=By.ID, value='btnSubmit').click()
            time.sleep(1)
        except:
            return None
        
        return driver
    
    def sieg_iris(driver):
        print('>>> Acessando IriS Dívida Ativa')
        try:
            driver.get('https://hub.sieg.com/IriS/#/DividaAtiva')
        except:
            return None
        
        return driver
    
    def consulta_lista(window_principal, driver, andamentos, pasta_final_download, pasta_final_dividas):
        ud = open(ultima_divida, 'r', encoding='utf-8').read()
        if ud != '':
            try:
                continuar_pagina = ud.split('|')[0]
                contador_dividas = int(ud.split('|')[1])
                cpf_cnpj_anterior = ud.split('|')[2]
                processo_anterior = ud.split('|')[3]
                inscricao_anterior = ud.split('|')[4]
                print(f'Última dívida salva:\n'
                      f'Página: {continuar_pagina}\n'
                      f'CPF/CNPJ: {cpf_cnpj_anterior}\n'
                      f'Número do processo: {processo_anterior}\n'
                      f'Número da inscrição: {inscricao_anterior}\n\n')
            
            except:
                continuar_pagina = ud
                contador_dividas = 0
                cpf_cnpj_anterior = 0
                processo_anterior = 0
                inscricao_anterior = 0
        else:
            continuar_pagina = 1
            contador_dividas = 0
            cpf_cnpj_anterior = 0
            processo_anterior = 0
            inscricao_anterior = 0
        
        print('>>> Consultando lista de arquivos\n')
        window_principal['-Mensagens-'].update('Aguardando lista de dívidas')
        window_principal.refresh()
        
        # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
        timer = 0
        while not comum.find_by_path('/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver
            time.sleep(1)
            timer += 1
            if timer >= 60:
                driver.close()
                return False
        
        window_principal['-Mensagens_2-'].update('')
        window_principal.refresh()
        time.sleep(2)
        if comum.find_by_id('modalYouTube', driver):
            # Encontre um elemento que esteja fora do modal e clique nele
            elemento_fora_modal = driver.find_element(by=By.ID, value='txtDataSearch')
            ActionChains(driver).move_to_element(elemento_fora_modal).click().perform()
        
        time.sleep(1)
        paginas = re.compile(r'>(\d+)</a></span><a class=\"paginate_button btn btn-default next').search(driver.page_source).group(1)
        
        contador = 1
        index = 0
        total = 0
        tempos = [datetime.now()]
        tempo_execucao = []
        for pagina in range(int(paginas) + 1):
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'STOP':
                return driver
            
            if pagina == 0:
                continue
            
            if pagina < int(continuar_pagina):
                window_principal['-Mensagens_2-'].update(f'Buscando última dívida salva, aguarde...')
                window_principal.refresh()
                driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
                # espera a lista de arquivos carregar
                while re.compile(r'id=\"tableActiveDebit_processing\" class=\"\" style=\"display: block;').search(driver.page_source):
                    time.sleep(1)
                time.sleep(3)
                continue
            
            # espera a lista de arquivos carregar
            while not comum.find_by_path('/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver
                
                time.sleep(1)
            
            time.sleep(1)
            print(f'>>> Página: {pagina}\n')
            with open(ultima_divida, 'w', encoding='utf-8') as f:
                f.write(str(pagina))
            
            info_pagina = driver.page_source
            
            detalhes = re.compile(r'btn-details float-right\" (data-id=\".+\") onclick=\"SeeDetails').findall(driver.page_source)
            print('Abrindo listas das empresas...')
            for detalhe in detalhes:
                pixels = 1
                while True:
                    try:
                        print(f'ID da lista: {detalhe}')
                        driver.find_element(by=By.XPATH, value=f'//a[@{detalhe}]').click()
                        time.sleep(0.1)
                        break
                    except:
                        driver.execute_script(f"window.scrollTo(0, {pixels})")
                        pixels += 1
                        pass
            
            time.sleep(1)
            # pega a lista de guias da competência desejada
            lista_divida = re.compile(r'excel\"><a id=\"(.+)\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"').findall(driver.page_source)
            
            nome = False
            # faz o download dos comprovantes
            for divida in lista_divida:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver
                
                for i in range(100):
                    id_empresa = re.compile(
                        r'<span class=\"title-datatable\">(.+)\((\d+)\)<span(.+\n){' + str(i) + '}.+col-md-2 td-background-dt\">(.+)</td><td class=\" td-background-dt\">(.+)</td><td class=\" td-background-dt hidden-sm hidden-xs\">(\d.+)</td><td class=\" td-background-dt hidden-sm hidden-xs\">.+excel\"><a id=\"' + divida).search(
                        driver.page_source)
                    if not id_empresa:
                        continue
                    nome = id_empresa.group(1).replace('&amp;', '&')
                    cpf_cnpj = id_empresa.group(2)
                    tipo = id_empresa.group(4)
                    processo = id_empresa.group(5)
                    inscricao = id_empresa.group(6)
                    break
                
                if not nome:
                    print(driver.page_source)
                    print(divida)
                    raise Exception('Rapadura é doce mas não é mole não')
                
                # verifica qual foi a última empresa consultada da lista
                if cpf_cnpj_anterior != 0:
                    # se achar a última dívida salva da execução anterior, pule mais um se não irá baixar a mesma dívida novamente
                    if not cpf_cnpj == cpf_cnpj_anterior and not processo == processo_anterior and not inscricao == inscricao_anterior:
                        window_principal['-Mensagens_2-'].update('Buscando última dívida salva, aguarde...')
                        window_principal.refresh()
                        continue
                    cpf_cnpj_anterior = 0
                    continue
                
                window_principal['-Mensagens_2-'].update('')
                window_principal.refresh()
                indices_dividas = re.compile(r'Mostrando (.+) até (.+) de (.+) </div>').search(driver.page_source)
                total_de_dividas  = int(indices_dividas.group(3).replace(',', '').replace('.', ''))
                if contador == 1:
                    index = contador_dividas
                    total = total_de_dividas-contador_dividas
                    
                # print(contador, 'fora do índice')
                tempos, tempo_execucao = comum.indice(contador, divida, total, index, window_principal, tempos, tempo_execucao)
                
                print('>>> Tentando baixar o arquivo')
                download = True
                while download:
                    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                    if cr == 'STOP':
                        return driver
                    
                    empresa = re.compile(r'col-md-3 td-background-dt hidden-xs\">(.+)</td><td class=\" td-background-dt hidden-xs\">.+excel\"><a id=\"' + divida + r'\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"')
                    try:
                        descricao = empresa.search(driver.page_source).group(1)
                    except:
                        descricao = 'SEM DESCRIÇÃO'
                    
                    driver, erro, download = download_divida(driver, divida, descricao, tipo, processo, inscricao, andamentos, pasta_final_download, pasta_final_dividas)
                    if erro == 'Finalizado':
                        return driver
                    # print(contador, 'depois do download')
                    if erro == 'erro':
                        try:
                            comum.escreve_relatorio_xlsx({'CNPJ': cpf_cnpj, 'NOME': nome, 'TIPO': tipo, 'PROCESSO': processo, 'INSCRIÇÃO': inscricao, 'RESULTADO': 'Erro ao baixar o arquivo'},
                                                         nome=andamentos, local=pasta_final)
                            print('>>> Erro ao baixar o arquivo\n')
                        except:
                            print(driver.page_source)
                            raise Exception(f'Então tá bom {divida}')
                        break
                
                contador += 1
                contador_dividas += 1
                with open(ultima_divida, 'w', encoding='utf-8') as f:
                    f.write(f'{pagina}|{contador_dividas}|{cpf_cnpj}|{processo}|{inscricao}')
                    
                
                
            if pagina == 25:
                return driver
            
            # próxima página
            try:
                driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
            except:
                return driver
            
            # espera a lista de arquivos carregar
            while re.compile(r'id=\"tableActiveDebit_processing\" class=\"\" style=\"display: block;').search(driver.page_source):
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver
                
                time.sleep(1)
            time.sleep(3)
            
            while info_pagina == driver.page_source:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'STOP':
                    return driver
                
                print(f'>>> Aguardando Página {pagina}\n')
                time.sleep(2)
            
            time.sleep(1)
        
        return driver
        
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=False, colunas_filtro=False, palavras_filtro=False)
    if continuar_rotina == '-reiniciar_rotina-':
        with open(ultima_divida, 'w', encoding='utf-8') as f:
            f.write('1')
    
    # captura o caminho absoluto do arquivo atual para criar uma pasta onde será feito o download dos arquivos para depois movê-los para a pasta definitiva
    if getattr(sys, 'frozen', False):
        # Se estiver em um executável criado por PyInstaller
        current_file_path = sys.executable
    else:
        # Se estiver em um script Python normal
        current_file_path = os.path.abspath(__file__)
    current_file_path_list = current_file_path.split('\\')[:-1]
    new_current_file_path = ''
    for caminho in current_file_path_list:
        new_current_file_path = new_current_file_path + caminho + '\\'
    pasta_final_download = new_current_file_path + 'Downloads'
    
    pasta_final_dividas = os.path.join(pasta_final, 'Dividas')
    # iniciar o driver do chome
    while True:
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            break
            
        status, driver = comum.configura_navegador(window_principal, pasta=pasta_final_download, timeout=60)
        driver = login_sieg(driver)
        if driver:
            driver = sieg_iris(driver)
            if driver:
                driver = consulta_lista(window_principal, driver, andamentos, pasta_final_download, pasta_final_dividas)
                if driver:
                    driver.close()
                    break

        window_principal['-Mensagens_2-'].update('SIEG IRIS demorou muito pra responder, tentando novamente')
        window_principal.refresh()
        print('SIEG IRIS demorou muito pra responder, tentando novamente')
    
    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
    if cr == 'STOP':
        alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
        return False, pasta_final
    
    return True, pasta_final


def run_pendencias_sigiss(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def login_sigiss(cnpj, senha):
        with Session() as s:
            # entra no site
            s.get('https://valinhos.sigissweb.com/')
            
            # loga na empresa
            query = {'loginacesso': cnpj,
                     'senha': senha}
            res = s.post('https://valinhos.sigissweb.com/ControleDeAcesso', data=query)
            
            try:
                soup = BeautifulSoup(res.content, 'html.parser')
                soup = soup.prettify()
                # print(soup)
                regex = re.compile(r"'Aviso', '(.+)<br>")
                regex2 = re.compile(r"'Aviso', '(.+)\.\.\.', ")
                try:
                    documento = regex.search(soup).group(1)
                except:
                    documento = regex2.search(soup).group(1)
                print(f"❌ {documento}")
                return False, documento, s
            except:
                print(f"✔")
                return True, '', s
    
    def consulta(s, cnpj, nome, pasta_final):
        print(f'>>> Consultando')
        s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=gerarcertidao')
        
        salvar = s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=imprimirCert&cnpjCpf=' + cnpj)
        # pego o contexto do link referente ao nome original do arquivo
        filename = salvar.headers.get('content-disposition', '')
        
        # aplico o regex para separar o nome do arquivo (Pendencias/Certidao)
        try:
            # crio um regex para obter o nome original do arquivo
            regex = re.compile(r'filename="(.*)\.pdf"')
            documento = regex.search(filename).group(1).replace('ê', 'e').replace('ã', 'a')
        except:
            # se não encontrar o nome do arquivo, procura por alguma mensagem de erro
            # pega o código da página
            soup = BeautifulSoup(salvar.content, 'html.parser')
            soup = soup.prettify()
            
            # procura no código da página a mensagem de erro
            mensagem = re.compile(r"mensagemDlg\((.+)',").search(soup).group(1)
            mensagem = mensagem.replace("','", ", ").replace("'", "")
            print(f"❌ {mensagem}")
            return False, mensagem, s
        
        # se for certidão cria uma pasta para a certidão
        print(f'>>> Salvando {documento}')
        if documento == 'Certidao':
            caminho = os.path.join(pasta_final, 'Certidões', cnpj + ' - ' + nome + ' - Certidão Negativa.pdf')
            os.makedirs(os.path.join(pasta_final, 'Certidões'), exist_ok=True)
            print(f"✔ Certidão")
        
        # se for pendência cria uma pasta para a pendência
        else:
            caminho = os.path.join(pasta_final, 'Pendências', cnpj + ' - ' + nome + ' - Pendências.pdf')
            os.makedirs(os.path.join(pasta_final, 'Pendências'), exist_ok=True)
            print(f"❗ Pendência")
        
        # pega a resposta da requisição que é o PDF codificado, e monta o arquivo.
        arquivo = open(caminho, 'wb')
        for parte in salvar.iter_content(100000):
            arquivo.write(parte)
        arquivo.close()
        
        return True, documento, s
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'Senha Prefeitura'], colunas_filtro=['Cidade'], palavras_filtro=['Valinhos'])
    if df_empresas.empty:
        return False, pasta_final
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, senha = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        print(cnpj)
        
        nome = nome.replace('/', ' ')
        
        # faz login no SIGISSWEB
        execucao, documento, s = login_sigiss(cnpj, senha)
        if execucao:
            # se fizer login, consulta a situação da empresa
            execucao, documento, s = consulta(s, cnpj, nome, pasta_final)
        
        # escreve os resultados da consulta
        comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'SENHA': senha, 'NOME':nome, 'RESULTADO':documento}, nome=andamentos, local=pasta_final)
        s.close()
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            return False, pasta_final
    
    return True, pasta_final


def executa_rotina(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final, rotina, continuar_rotina):
    rotinas = {'Consulta Certidão Negativa de Débitos Tributários Não Inscritos': run_cndtni,
               'Consulta Débitos Municipais Jundiaí': run_debitos_municipais_jundiai,
               'Consulta Débitos Municipais Valinhos': run_debitos_municipais_valinhos,
               'Consulta Débitos Municipais Vinhedo': run_debitos_municipais_vinhedo,
               'Consulta Débitos Estaduais - Situação do Contribuinte': run_debitos_estaduais,
               'Consulta Dívida Ativa da União': run_debitos_divida_ativa_uniao,
               'Consulta Dívida Ativa Procuradoria Geral do Estado de São Paulo': run_debitos_divida_ativa,
               'Consulta Pendências SIGISSWEB Valinhos': run_pendencias_sigiss}
    
    tempos = [datetime.now()]
    tempo_execucao = []
    # try:
    resultado, pasta_final_oficial = rotinas[str(rotina)](window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final, rotina, tempos, tempo_execucao, continuar_rotina)
    """except ValueError:
        alert(f'❌ A planilha de dados selecionada "{planilha_dados}" não contem as colunas necessárias para a consulta')
        return"""
        
    if resultado:
        window_principal['-Mensagens-'].update('')
        window_principal['-Mensagens_2-'].update('')
        window_principal['-progressbar-'].update_bar(100, max=100)
        window_principal['-Progresso_texto-'].update('100 %')
        alert(f'✔ Rotina "{rotina}" finalizada')
    else:
        try:
            os.remove(os.path.join(pasta_final_oficial, 'Dados.xlsx'))
            os.rmdir(pasta_final_oficial)
        except:
            pass
