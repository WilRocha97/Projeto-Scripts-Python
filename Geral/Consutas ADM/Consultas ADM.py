# -*- coding: utf-8 -*-
import chromedriver_autoinstaller, fitz, shutil, time, io, re, os, sys, traceback, pandas as pd, PySimpleGUI as sg
from datetime import datetime, timedelta
from win32com import client
from requests import Session
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from anticaptchaofficial.imagecaptcha import *
from anticaptchaofficial.recaptchav2proxyless import *
from anticaptchaofficial.recaptchav3proxyless import *
from anticaptchaofficial.hcaptchaproxyless import *
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from PIL import Image
from threading import Thread
from pyautogui import confirm, alert

icone = 'Assets/auto-flash.ico'
dados_anticaptcha = "Dados\Dados AC.txt"
dados_contadores = "Dados\Contadores.txt"
dados_modo = 'Assets/modo.txt'

def concatena(variavel, quantidade, posicao, caractere):
    # função para concatenar strings colocando caracteres no começo ou no final
    variavel = str(variavel)
    if posicao == 'depois':
        # concatena depois
        while len(str(variavel)) < quantidade:
            variavel += str(caractere)
    if posicao == 'antes':
        # concatena antes
        while len(str(variavel)) < quantidade:
            variavel = str(caractere) + str(variavel)
    
    return variavel


def initialize_chrome(options=webdriver.ChromeOptions()):
    # biblioteca para baixar o chormedriver atualizado
    chromedriver_autoinstaller.install()
    print('>>> Inicializando Chromedriver...')
    
    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
    
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # retorna o chormedriver aberto
    return True, webdriver.Chrome(options=options)


def find_by_id(item, driver):
    try:
        driver.find_element(by=By.ID, value=item)
        return True
    except:
        return False


def find_by_path(item, driver):
    try:
        driver.find_element(by=By.XPATH, value=item)
        return True
    except:
        return False
    

def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass


def send_input_xpath(elem_path, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_path)
            elem.send_keys(data)
            break
        except:
            pass


def get_info_post(soup):
    # captura infos para realizar um post
    soup = BeautifulSoup(soup, 'html.parser')
    infos = [
        soup.find('input', attrs={'id': '__VIEWSTATE'}),
        soup.find('input', attrs={'id': '__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id': '__EVENTVALIDATION'})
    ]
    
    # state, generator, validation
    return tuple(info.get('value', '') for info in infos if info)


def new_session_fazenda_driver(window_principal, user, pwd, perfil, retorna_driver=False, options='padrão'):
    url_home = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"
    _site_key = '6LesbbcZAAAAADrEtLsDUDO512eIVMtXNU_mVmUr'
    
    window_principal['-Mensagens_2-'].update('Iniciando ambiente da consulta, aguarde...')
    window_principal.refresh()
    
    if options == 'padrão':
        # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--window-size=1920,1080')
        # options.add_argument("--start-maximized")
    
    status, driver = initialize_chrome(options)
    driver.get(url_home)
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': _site_key, 'url': url_home}
    token = solve_recaptcha(recaptcha_data)
    token = str(token)
    
    if event == '-encerrar-' or event == sg.WIN_CLOSED:
        return False, ''
    
    if perfil == 'contador':
        window_principal['-Mensagens_2-'].update(f'Abrindo perfil do contador...')
        window_principal.refresh()
        button = driver.find_element(by=By.ID, value='ConteudoPagina_rdoListPerfil_1')
        button.click()
        time.sleep(1)
    
    elif perfil == 'contribuinte':
        window_principal['-Mensagens_2-'].update(f'Abrindo perfil do contribuinte...')
        window_principal.refresh()
    
    print(f'>>> Logando no usuário')
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtUsuario')
    element.send_keys(user)
    
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtSenha')
    element.send_keys(pwd)
    
    script = 'document.getElementById("g-recaptcha-response").innerHTML="{}";'.format(token)
    driver.execute_script(script)
    
    script = 'document.getElementById("ConteudoPagina_btnAcessar").disabled = false;'
    driver.execute_script(script)
    time.sleep(1)
    
    button = driver.find_element(by=By.ID, value='ConteudoPagina_btnAcessar')
    button.click()
    time.sleep(3)
    
    try:
        button = driver.find_element(by=By.XPATH, value='/html/body/div[2]/section/div/div/div/div[2]/div/ul/li/form/div[5]/div/a')
        button.click()
        time.sleep(3)
    except:
        pass
    
    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    padrao = re.compile(r'SID=(.\d+)')
    resposta = padrao.search(str(soup))
    
    if not resposta:
        try:
            padrao = re.compile(r'(Senha inserida está incorreta)')
            resposta = padrao.search(str(soup))
            
            if not resposta:
                padrao = re.compile(r'(Favor informar o login e a senha corretamente.)')
                resposta = padrao.search(str(soup))
                
                if not resposta:
                    padrao = re.compile(r'(O usuário não tem acesso a este serviço do sistema ou o serviço não foi liberado para a sua utilização.)')
                    resposta = padrao.search(str(soup))
                    
                    if not resposta:
                        padrao = re.compile(r'(ERRO INTERNO AO SISTEMA DE CONTROLE DE ACESSO)')
                        driver.save_screenshot(r'ignore\debug_screen.png')
                        resposta = padrao.search(str(soup))
            
            sid = resposta.group(1)
            print(f'❌ {sid}')
            cokkies = 'erro'
            driver.close()
            
            return cokkies, sid
        except:
            driver.save_screenshot(r'Log\debug_screen.png')
            driver.close()
            return False, 'Erro ao logar no perfil'
    
    if retorna_driver:
        sid = resposta.group(1)
        return driver, sid
    
    sid = resposta.group(1)
    cookies = driver.get_cookies()
    driver.quit()
    
    return cookies, sid


def solve_recaptcha(data):
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_anti = f.read()
    anticaptcha_api_key = key_anti
    
    print('>>> Quebrando recaptcha')
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    solver.set_website_url(data['url'])
    solver.set_website_key(data['sitekey'])
    # set optional custom parameter which Google made for their search page Recaptcha v2
    # solver.set_data_s('"data-s" token from Google Search results "protection"')

    g_response = solver.solve_and_return_solution()
    if g_response != 0:
        return g_response
    else:
        print(solver.error_code)
        return solver.error_code


def solve_text_captcha(driver, captcha_element, element_type='id'):
    element = ''
    
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_anti = f.read()
    anticaptcha_api_key = key_anti
    
    os.makedirs('Log\captcha', exist_ok=True)
    # captura a imagem do captcha
    if element_type == 'id':
        element = driver.find_element(by=By.ID, value=captcha_element)
    elif element_type == 'xpath':
        element = driver.find_element(by=By.XPATH, value=captcha_element)
    
    location = element.location
    size = element.size
    driver.save_screenshot('Log\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    time.sleep(2)
    im = Image.open(r'Log\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'Log\captcha\captcha.png')
    time.sleep(1)
    
    print('>>> Quebrando text captcha')
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    
    captcha_text = solver.solve_and_return_solution(os.path.join('ignore', 'captcha', 'captcha.png'))
    if captcha_text != 0:
        return captcha_text
    else:
        print(solver.error_code)
        return solver.error_code


def indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao):
    tempo_estimado = 0
    
    # captura a hora atual e coloca em uma lista para calcular o tempo de execução do andamento atual
    tempo_inicial = datetime.now()
    tempos.append(tempo_inicial)
    tempo_execucao_atual = int(tempos[1].timestamp()) - int(tempos[0].timestamp())
    
    # adiciona o tempo de execução atual na lista com os tempos anteriores para calcular a média de tempo de execução dos andamentos
    tempo_execucao.append(tempo_execucao_atual)
    for t in tempo_execucao:
        tempo_estimado = tempo_estimado + t
    tempo_estimado = int(tempo_estimado) / int(len(tempo_execucao))
    
    # multiplica o tempo médio de execução dos andamentos pelo número de andamentos que faltam executar para obter o tempo estimado em segundos
    tempo_total_segundos = int((len(total_empresas) + index) - (count + index) + 1) * float(tempo_estimado)
    # Converter o tempo total para um objeto timedelta
    tempo_total = timedelta(seconds=tempo_total_segundos)
    
    # Extrair dias, horas e minutos do timedelta
    dias = tempo_total.days
    horas = tempo_total.seconds // 3600
    minutos = (tempo_total.seconds % 3600) // 60
    segundos = (tempo_total.seconds % 3600) % 60
    
    # Retorna o tempo no formato "dias:horas:minutos:segundos"
    dias_texto = ''
    horas_texto = ''
    minutos_texto = ''
    segundos_texto = ''
    
    if dias > 0:
        dias_texto = f' {dias} dias'
    if horas > 0:
        horas_texto = f' {horas} horas'
    if minutos > 0:
        minutos_texto = f' {minutos} minutos'
    if segundos > 0:
        segundos_texto = f' {segundos} segundos'
    
    tempo_estimado_texto = f" | Tempo estimado:{dias_texto}{horas_texto}{minutos_texto}{segundos_texto}"
    
    tempos.pop(0)
    
    print(f'\n\n[{empresa}]')
    window_principal['-Mensagens-'].update(f'{str((count + index) - 1)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index) + 1)} Restantes{tempo_estimado_texto}')
    window_principal['-progressbar-'].update_bar(count, max=int(len(total_empresas)))
    window_principal['-Progresso_texto-'].update(str(round(float(count) / int(len(total_empresas)) * 100, 1)) + '%')
    window_principal.refresh()
    
    tempo_estimado = tempo_execucao
    return tempos, tempo_estimado


def open_dados(andamentos, codigo_20000, pasta_final, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, filtrar_celulas_em_branco):
    def open_lista_dados(dados_final, encode):
        try:
            with open(dados_final, 'r', encoding=encode) as f:
                dados = f.readlines()
        except Exception as e:
            alert(f'❌ Não pode abrir arquivo\n{planilha_dados}\n{str(e)}')
            return False
        
        return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))
    
    dados_final = os.path.join(pasta_final, 'Dados.csv')
    encode = 'latin-1'
    
    # modelo de lista com as colunas que serão usadas na rotina
    # colunas_usadas = ['column1', 'column2', 'column3']
    
    df = pd.read_excel(planilha_dados)
    
    # coluna com os códigos do ae
    coluna_codigo = 'Codigo'
    
    if codigo_20000 == '-codigo_20000_nao-':
        empresas_20000 = 'Empresas com o código menor que 20.000'
        # cria um novo df apenas com empresas a baixo do código 20.000
        df_filtrada = df[df[coluna_codigo] <= 20000]
    elif codigo_20000 == '-codigo_20000-':
        empresas_20000 = 'Empresas com o código maior que 20.000'
        # cria um novo df apenas com empresas a cima do código 20.000
        df_filtrada = df[df[coluna_codigo] >= 20000]
    else:
        empresas_20000 = 'Empresas com qualquer código'
        df_filtrada = df
    
    # filtra as células de colunas específicas que contenham palavras especificas
    if palavras_filtro and colunas_filtro:
        for count, coluna_para_filtrar in enumerate(colunas_filtro):
            df_filtrada = df_filtrada[df_filtrada[coluna_para_filtrar].str.contains(palavras_filtro[count], case=False, na=False)]
    
    # filtra as colunas
    try:
        df_filtrada = df_filtrada[colunas_usadas]
    except KeyError:
        alert(f'❌ Erro ao buscar as colunas na planilha base selecionada: {planilha_dados}\n\n'
                              f'Verifique se a planilha contem as colunas necessárias para a execução da rotina e se elas tem exatamente o mesmo nome indicado ao lado: {colunas_usadas}')
        return False
    
    if filtrar_celulas_em_branco:
        df_filtrada = df_filtrada.dropna(subset=filtrar_celulas_em_branco)
        # df_filtrada = df_filtrada.fillna('vazio')
    else:
        # remove linha com células vazias
        df_filtrada = df_filtrada.dropna(axis=0, how='any')
    
    # Converte a coluna 'CNPJ' para string e remova a parte decimal '.0'. Preencha com zeros à esquerda para garantir 14 dígitos
    df_filtrada['CNPJ'] = df_filtrada['CNPJ'].astype(str).str.replace(r'\.0', '', regex=True).str.zfill(14)
    
    if andamentos == 'Consulta Débitos Estaduais - Situação do Contribuinte' or andamentos == 'Consulta Certidão Negativa de Débitos Tributários Não Inscritos':
        contadores_dict = atualiza_contadores()
        # Substituir valores com base no dicionário apenas se o valor estiver presente no dicionário
        df_filtrada['Perfil'] = 'vazio'
        
        # Função para atualizar os valores das colunas com base no dicionário de mapeamento
        def atualizar_valores(row):
            if row['PostoFiscalContador'] in contadores_dict:
                return contadores_dict[row['PostoFiscalContador']]
            else:
                return (row['PostoFiscalUsuario'], row['PostoFiscalSenha'], 'contribuinte')
        
        # Aplicar a função para atualizar os valores das colunas
        df_filtrada[['PostoFiscalUsuario', 'PostoFiscalSenha', 'Perfil']] = df_filtrada.apply(atualizar_valores, axis=1, result_type='expand')
    
        
        # 5. Deletar a coluna 'contador'
        df_filtrada.drop(columns=['PostoFiscalContador'], inplace=True)
        
        # 3. Deletar linhas com células vazias na coluna 'senha'
        df_filtrada = df_filtrada.dropna(subset=['PostoFiscalSenha'])
        
        # Ordene o DataFrame com base na coluna desejada
        df_filtrada = df_filtrada.sort_values(by=['Perfil', 'PostoFiscalUsuario', 'CNPJ'], ascending=[True, True, True])
        
        # remove linha com células vazias
        df_filtrada = df_filtrada.dropna(axis=0, how='any')
        
        # Remover linhas que contenham 'ISENTO' na coluna 'PostoFiscalUsuario'
        df_filtrada = df_filtrada[~df_filtrada['PostoFiscalUsuario'].str.contains('ISENTO', case=False, na=False)]
        # Remover linhas que contenham 'BAIXADO' na coluna 'PostoFiscalUsuario'
        df_filtrada = df_filtrada[~df_filtrada['PostoFiscalUsuario'].str.contains('BAIXADO', case=False, na=False)]
        
    if df_filtrada.empty:
        alert(f'❗ Não foi encontrado nenhuma empresa na planilha selecionada: {planilha_dados}\n\n'
                               f'{empresas_20000}\n'
                               f'utilizando os seguintes filtros: {palavras_filtro}\n'
                               f'nas respectivas colunas {colunas_filtro}\n')
        return False
    
    for coluna in df_filtrada.columns:
        # Remova aspas duplas
        df_filtrada[coluna] = df_filtrada[coluna].str.replace('"', '')
        
        # Remova quebras de linha (`\n` e `\r`)
        df_filtrada[coluna] = df_filtrada[coluna].str.replace('\n', '').str.replace('\r', '').str.replace('_x000D_', '')
    
    df_filtrada.to_csv(dados_final, header=False, index=False, sep=';', encoding=encode)
    
    empresas = open_lista_dados(dados_final, encode)
    return empresas


def escreve_doc(texto, local='Log', nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    try:
        os.remove(os.path.join(local, 'Log.txt'))
    except:
        pass
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()


def escreve_relatorio_csv(texto, local, nome='Relatório', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    contador = 0
    while True:
        try:
            f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
            break
        except:
            contador += 1
            time.sleep(1)
            if contador > 30:
                try:
                    f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
                    break
                except:
                    try:
                        f = open(os.path.join(local, f"{nome} - Parte 2.csv"), 'a', encoding=encode)
                        break
                    except:
                        f = open(os.path.join(local, f"{nome} - Parte 3.csv"), 'a', encoding=encode)
                        break
            pass
        
    f.write(texto + '\n')
    f.close()
    
    
def configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos, colunas_usadas=None, colunas_filtro=None, palavras_filtro=None, filtrar_celulas_em_branco=None):
    def where_to_start(idents, pasta_final_anterior, planilha, encode='latin-1'):
        if not os.path.isdir(pasta_final_anterior):
            return 0
        
        file = os.path.join(pasta_final_anterior, planilha)
        
        try:
            with open(file, 'r', encoding=encode) as f:
                dados = f.readlines()
        except:
            alert(f'❗ Não foi encontrada nenhuma planilha de andamentos na pasta de execução anterior.\n\n'
                  f'Começando a execução a partir do primeiro indice da planilha de dados selecionada.')
            return 0
        
        # Busca o último andamento, esse 'for' é necessário poís pode acontecer da planilha anterior ser editada e ficar linha em branco no final dela.
        # Busca os últimos indices de baixo para cima: -1 -2 -3... até -10000
        for i in range(-1, -10001, -1):
            try:
                elem = dados[i].split(';')[0]
                if len(idents) == idents.index(elem) + 1:
                    return 0
                return idents.index(elem) + 1
            except ValueError:
                continue
        return 0
    
    comp = datetime.now().strftime('%m-%Y')
    pasta_final_ = os.path.join(pasta_final_, andamentos, comp)
    contador = 0
    
    # iteração para determinar se precisa criar uma pasta nova para armazenar os resultados
    # toda vês que o programa começar as consultas uma nova pasta será criada para não sobrepor ou misturar as execuções
    while True:
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            return '', '', False
        try:
            os.makedirs(os.path.join(pasta_final_, 'Execuções'))
            pasta_final = os.path.join(pasta_final_, 'Execuções')
            pasta_final_anterior = False
            break
        except:
            try:
                contador += 1
                os.makedirs(os.path.join(pasta_final_, f'Execuções ({str(contador)})'))
                pasta_final = os.path.join(pasta_final_, f'Execuções ({str(contador)})')
                if contador - 1 < 1:
                    pasta = 'Execuções'
                else:
                    pasta = f'Execuções ({str(contador - 1)})'
                pasta_final_anterior = os.path.join(pasta_final_, pasta)
                break
            except:
                pass
    
    window_principal['-Mensagens-'].update('Criando dados para a consulta...')
    window_principal.refresh()
    # abrir a planilha de dados
    empresas = open_dados(andamentos, codigo_20000, pasta_final, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, filtrar_celulas_em_branco)
    if not empresas:
        return '', '', False
    
    if pasta_final_anterior:
        planilha = f'{andamentos}.csv'
        # obtêm o indice do último andamento da execução anterior para continuar
        index = where_to_start(tuple(i[0] for i in empresas), pasta_final_anterior, planilha)
    else:
        index = 0
    
    return pasta_final, index, empresas


def atualiza_contadores():
    # obtêm a lista de usuário e senha de cada contador para a planilha dedados de algumas consultas
    f = open(dados_contadores, 'r', encoding='utf-8')
    contadores = f.readlines()
    
    contadores_dict = {}
    for contador in contadores:
        contador = contador.split('/')
        contadores_dict[contador[0]] = (contador[1],contador[2],contador[3])
    
    return contadores_dict


def download_file(name, response, pasta):
    # função para salvar um arquivo retornado de uma requisição
    pasta = str(pasta).replace('\\', '/')
    os.makedirs(pasta, exist_ok=True)

    with open(os.path.join(pasta, name), 'wb') as arq:
        for i in response.iter_content(100000):
            arq.write(i)
            

# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ FUNÇÕES EM COMUM ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ ROTINAS ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓


def run_cndtni(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
    def abre_pagina_consulta(window_principal, driver):
        print('>>> Abrindo Conta Fiscal')
        window_principal['-Mensagens_2-'].update(f'Abrindo Conta Fiscal...')
        window_principal.refresh()
        
        try:
            # iteração para logar no usuário
            while not re.compile(r'>Conta Fiscal do ICMS e Parcelamento').search(driver.page_source):
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
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
        
        while not find_by_id('divcontainer', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                return False, ''
            time.sleep(1)
        
        print('>>> Abrindo consulta de CNDNI')
        window_principal['-Mensagens_2-'].update(f'Abrindo consulta de CNDNI...')
        window_principal.refresh()
        # iteração para capturar a url da página final para poder inserir os infos da empresa e consultar
        while True:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                return False, ''
            try:
                url_consulta_cndni = re.compile(r'href="(https:\/\/www10.fazenda.sp.gov.br\/CertidaoNegativaDeb\/Pages.+)" tabindex="-1">Verificar Impedimentos eCND').search(driver.page_source).group(1)
                break
            except:
                pass
        
        # abre a página
        driver.get(url_consulta_cndni)
        
        return driver, 'ok'
    
    def consulta_cndni(driver, nome, cnpj, pasta_download, pasta_final):
        print('>>> Consultando.')
        while not find_by_id('MainContent_txtDocumento', driver):
            time.sleep(1)
        
        # enquanto a tela com os resultados não abre, tenta logar na empresa, se der erro de CNPJ tenta novamente mais duas vezes, se não conseguir retorna erro de CNPJ
        # fora esse erro tenta logar no usuário mais 15 vezes no total sempre verificando se algúm alerta for exibido, caso seja retorna a mensagem do alerta
        contador = 0
        contador_cpf = 0
        while not find_by_id('MainContent_lnkImprimirCertidaoBotao1', driver):
            while True:
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
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
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
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
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
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
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    return False
                try:
                    doc = fitz.open(arq, filetype="pdf")
                    break
                except:
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
            # move o arquivo com novo nome para melhor identifica-lo
            shutil.move(arq, os.path.join(pasta_sem_debito, f'{nome[:30]} - {cnpj} - CNDNI.pdf'))
        
        print('✔ Sem débitos')
        return 'Sem débitos'
    
    # bloco para filtrar e criar a nova planilha de dados
    colunas_usadas = ['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador']
    filtrar_celulas_em_branco = ['CNPJ', 'Razao', 'Cidade']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, filtrar_celulas_em_branco=filtrar_celulas_em_branco)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
        
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    driver = ''
    
    pasta_download = os.path.join(pasta_final, 'Download')
    
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": pasta_download.replace('/', '\\'),  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = concatena(cnpj, 14, 'antes', 0)
        nome = nome.replace('/', '')
        print(cnpj)
        
        resultado = 'ok'
        while True:
            # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
            if usuario_anterior != usuario:
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
                try:
                    driver.close()
                except:
                    pass
                
                contador = 0
                while True:
                    if event == '-encerrar-' or event == sg.WIN_CLOSED:
                        alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                        try:
                            driver.close()
                        except:
                            pass
                        try:
                            os.rmdir(pasta_final)
                        except:
                            pass
                        return False
                    try:
                        # abre uma nova sessão no site da fazenda
                        driver, sid = new_session_fazenda_driver(window_principal, usuario, senha, perfil, retorna_driver=True, options=options)
                        break
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
                    
                    if contador >= 5:
                        print('❌ Impossível de logar com esse usuário')
                        sid = 'Impossível de logar com esse usuário'
                        driver = 'erro'
                        break
                
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                if driver == 'erro':
                    escreve_relatorio_csv(f'{cnpj};{nome};{cidade};{sid}', nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                else:
                    driver, resultado = abre_pagina_consulta(window_principal, driver)
            
            # se o resultado da abertura da página de consulta for 'ok', consulta a empresa
            if resultado == 'ok':
                driver, resultado = consulta_cndni(driver, nome, cnpj, pasta_download, pasta_final)
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                if resultado != 'erro':
                    escreve_relatorio_csv(f'{cnpj};{nome};{cidade};{resultado}', nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                driver.close()
                usuario_anterior = 'padrão'
                continue
                
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        driver.close()
                    except:
                        pass
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
    
    try:
        driver.close()
    except:
        pass
    return True

    
def run_debitos_municipais_jundiai(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
    def login_jundiai(window_principal, cnpj, insc_muni, pasta_final):
        cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
        
        pasta_final_certidao = os.path.join(pasta_final, 'Certidões')
        
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        # options.add_argument('--window-size=1920,1080')
        # options.add_argument("--start-maximized")
        options.add_experimental_option('prefs', {
            "download.default_directory": pasta_final_certidao.replace('/', '\\'),  # Change default directory for downloads
            "download.prompt_for_download": False,  # To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
        })
        status, driver = initialize_chrome(options)
        
        # entra na página inicial da consulta
        url_inicio = 'https://web.jundiai.sp.gov.br/PMJ/SW/certidaonegativamobiliario.aspx'
        driver.get(url_inicio)
        time.sleep(1)
        
        # pega o sitekey para quebrar o captcha
        data = {'url': url_inicio, 'sitekey': '6LfK0igTAAAAAOeUqc7uHQpW4XS3EqxwOUCHaSSi'}
        response = solve_recaptcha(data)
        
        # insere a inscrição estadual e o cnpj da empresa
        send_input('DadoContribuinteMobiliario1_txtCfm', insc_muni, driver)
        send_input('DadoContribuinteMobiliario1_txtNrCicCgc', cnpj, driver)
        
        # pega a id do campo do recaptcha
        id_response = re.compile(r'class=\"recaptcha\".+<textarea id=\"(.+)\" name=')
        id_response = id_response.search(driver.page_source).group(1)
        
        text_response = ''
        while not text_response == response:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False, ''
            
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
        
        while not find_by_id('lblContribuinte', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False, ''
            
            if find_by_id('AjaxAlertMod1_lblAjaxAlertMensagem', driver):
                situacao = re.compile(r'AjaxAlertMod1_lblAjaxAlertMensagem\">(.+)</span>')
                while True:
                    if event == '-encerrar-' or event == sg.WIN_CLOSED:
                        try:
                            driver.close()
                        except:
                            pass
                        
                        try:
                            os.rmdir(pasta_final)
                        except:
                            pass
                        return False, ''
                    
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
                    driver.close()
                    return situacao, situacao_print
                if situacao == 'Confirme que você não é um robô':
                    situacao_print = '❌ Erro ao logar na empresa'
                    driver.close()
                    return 'Erro ao logar na empresa', situacao_print
                if situacao == 'EIX000: (-107) ISAM error:  record is locked.':
                    situacao_print = '❌ Erro no site'
                    driver.close()
                    return 'Erro no site', situacao_print
            time.sleep(1)
        
        situacao = re.compile(r'<span id=\"lblSituacao\">(.+)</span>')
        situacao = situacao.search(driver.page_source).group(1)
        time.sleep(1)
        
        if situacao == 'Não constam débitos para o contribuinte':
            os.makedirs(pasta_final_certidao, exist_ok=True)
            
            send_input('txtSolicitante', 'Escritório', driver)
            send_input('txtMotivo', 'Consulta', driver)
            time.sleep(1)
            driver.find_element(by=By.ID, value='btnImprimir').click()
            
            while not os.path.exists(os.path.join(pasta_final_certidao, 'relatorio.pdf')):
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    try:
                        driver.close()
                    except:
                        pass
                    
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False, ''
                time.sleep(1)
            while os.path.exists(os.path.join(pasta_final_certidao, 'relatorio.pdf')):
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    try:
                        driver.close()
                    except:
                        pass
                    
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False, ''
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
            driver.quit()
            return situacao, situacao_print
        
        situacao_print = f'✔ {situacao}'
        driver.quit()
        return situacao, situacao_print
    
    colunas_usadas = ['CNPJ', 'InsMunicipal', 'Razao']
    colunas_filtro = ['Cidade']
    palavras_filtro = ['Jundiaí']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, colunas_filtro=colunas_filtro, palavras_filtro=palavras_filtro)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome = empresa
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        situacao = ''
        situacao_print = ''
        contador = 0
        # iteração para logar na empresa e consultar, tenta 5 vezes
        while True:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                break
            
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            situacao, situacao_print = login_jundiai(window_principal, cnpj, insc_muni, pasta_final)
            if situacao == 'Erro ao logar na empresa':
                contador += 1
                if contador >= 5:
                    break
                continue
            
            if situacao != 'Erro no site':
                break
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            try:
                driver.close()
            except:
                pass
            
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return False
        
        print(situacao_print)
        escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{situacao}', nome=andamentos, local=pasta_final)
    
    return True


def run_debitos_municipais_valinhos(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
        while not find_by_id('span7Menu', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        button = driver.find_element(by=By.ID, value='span7Menu')
        button.click()
        
        while not find_by_id('captchaimg', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        
        time.sleep(1)
        captcha = solve_text_captcha(driver, 'captchaimg')
        
        if not captcha:
            print('Erro Login - não encontrou captcha')
            return driver, 'erro captcha'
        
        while not find_by_id('input1', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input1')
        element.send_keys(insc_muni)
        
        while not find_by_id('input4', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='input4')
        element.send_keys(cnpj)
        
        while not find_by_id('captchafield', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        element = driver.find_element(by=By.ID, value='captchafield')
        element.send_keys(captcha)
        
        while not find_by_id('imagebutton1', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            time.sleep(1)
        button = driver.find_element(by=By.ID, value='imagebutton1')
        button.click()
        
        time.sleep(3)
        
        timer = 0
        while not find_by_id('td30', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                return
            print('>>> Aguardando site')
            driver.save_screenshot(r'Log\debug_screen.png')
            time.sleep(1)
            
            if find_by_id('tdMsg', driver):
                try:
                    erro = re.compile(r'id=\"tdMsg\".+tipo=\"td\">(.+)\.').search(driver.page_source).group(1)
                    print(f'❌ {erro}')
                    return driver, erro
                except:
                    pass
            timer += 1
            if timer >= 10:
                print(f'❌ Erro no login')
                return driver, 'Erro no login'
        
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
                    pisa.showLogging()
                    pisa.CreatePDF(str_html, pdf)
                    print('❗ Arquivo gerado')
                escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};Com débitos', nome=andamentos, local=pasta_final)
            else:
                escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};Sem débitos', nome=andamentos, local=pasta_final)
                print('✔ Não há débitos')
        except:
            escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};Erro na geração do PDF', nome=andamentos, local=pasta_final)
            driver.save_screenshot(r'Log\debug_screen.png')
            print('❌ Erro na geração do PDF')
        
        return driver, True
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    colunas_usadas = ['CNPJ', 'InsMunicipal', 'Razao']
    colunas_filtro = ['Cidade']
    palavras_filtro = ['Valinhos']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, colunas_filtro=colunas_filtro, palavras_filtro=palavras_filtro)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome = empresa
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        try:
            insc_muni = int(insc_muni)
        except:
            continue
        
        resultado = 'Texto da imagem incorreto'
        while resultado == 'Texto da imagem incorreto':
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False
            
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            
            status, driver = initialize_chrome(options)
            driver, resultado = login_valinhos(driver, cnpj, insc_muni)
            
            if resultado == 'ok':
                driver, resultado = consulta_valinhos(driver, cnpj, insc_muni, nome, andamentos, pasta_final)
            elif resultado == 'Texto da imagem incorreto':
                pass
            else:
                escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{resultado}', nome=andamentos, local=pasta_final)
            driver.close()
            
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return False
    
    return True
    
   
def run_debitos_municipais_vinhedo(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
    def pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final):
        # opções para fazer com que o chome trabalhe em segundo plano (opcional)
        options = webdriver.ChromeOptions()
        # options.add_argument('--headless')
        # options.add_argument('--window-size=1920,1080')
        options.add_argument("--start-maximized")
        options.add_experimental_option('prefs', {
            "download.default_directory": pasta_final.replace('/', '\\'),  # Change default directory for downloads
            "download.prompt_for_download": False,  # To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
        })
        status, driver = initialize_chrome(options)
        
        url_inicio = 'http://vinhedomun.presconinformatica.com.br/certidaoNegativa.jsf?faces-redirect=true'
        driver.get(url_inicio)
        
        contador = 1
        while not find_by_id(f'homeForm:panelCaptcha:j_idt{str(contador)}', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return
            contador += 1
            time.sleep(0.2)
        
        # resolve o captcha
        captcha = solve_text_captcha(driver, f'homeForm:panelCaptcha:j_idt{str(contador)}')
        
        # espera o campo do tipo da pesquisa
        while not find_by_id('homeForm:inputTipoInscricao_label', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return
            time.sleep(1)
        # clica no menu
        driver.find_element(by=By.ID, value='homeForm:inputTipoInscricao_label').click()
        
        # espera o menu abrir
        while not find_by_path('/html/body/div[6]/div/ul/li[2]', driver):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return
            time.sleep(1)
        # clica na opção "Mobiliário"
        driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/ul/li[2]').click()
        
        if not captcha:
            print('Erro Login - não encontrou captcha')
            return False, 'recomeçar'
        
        time.sleep(1)
        # clica na barra de inscrição e insere
        driver.find_element(by=By.ID, value='homeForm:inputInscricao').click()
        time.sleep(2)
        driver.find_element(by=By.ID, value='homeForm:inputInscricao').send_keys(insc_muni)
        send_input(f'homeForm:panelCaptcha:j_idt{str(contador + 3)}', captcha, driver)
        time.sleep(2)
        
        # clica no botão de pesquisar
        driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[2]/div[5]/a/div').click()
        
        print('>>> Consultando')
        window_principal['-Mensagens_2-'].update('Consultando empresa...')
        window_principal.refresh()
        while 'dados-contribuinte-inscricao">0000000000' not in driver.page_source:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return
            
            if find_by_path('/html/body/div[1]/div[5]/div[2]/div[1]/div/ul/li/span', driver):
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
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            try:
                driver.close()
            except:
                pass
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return
        
        situacao_print = f'✔ {situacao}'
        return situacao, situacao_print
    
    def salvar_guia_vinhedo(window_principal, driver, cnpj, pasta_final):
        window_principal['-Mensagens_2-'].update('Buscando Certidão Negativa...')
        window_principal.refresh()
        time.sleep(1)
        while True:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
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
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                try:
                    driver.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
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
    
    colunas_usadas = ['CNPJ', 'InsMunicipal', 'Razao']
    colunas_filtro = ['Cidade']
    palavras_filtro = ['Vinhedo']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, colunas_filtro=colunas_filtro, palavras_filtro=palavras_filtro)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)

        cnpj, insc_muni, nome = empresa
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        while True:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False
            
            window_principal['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_principal.refresh()
            
            situacao, situacao_print = pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final)
            if situacao_print != 'recomeçar':
                print(situacao_print)
                if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
                    alert(f'❗ Rotina "{andamentos}":\n\nDesculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.')
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{situacao}', nome=andamentos, local=pasta_final)
                break
    
    return True


def run_debitos_estaduais(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                return False
            try:
                res = s.get(url_pesquisa)
                time.sleep(2)
                
                state, generator, validation = get_info_post(res.content)
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
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                return False
            
            try:
                res = s.post(url_pesquisa, info)
                soup = BeautifulSoup(res.content, 'html.parser')
                
                state, generator, validation = get_info_post(res.content)
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
            download_file(nome, impressao, pasta=pasta_final)
        else:
            situacao = situacoes_debitos_estaduais['E']
        
        return situacao
    
    colunas_usadas = ['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador']
    filtrar_celulas_em_branco = ['CNPJ', 'Razao', 'Cidade']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, filtrar_celulas_em_branco=filtrar_celulas_em_branco)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    s = False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = concatena(cnpj, 14, 'antes', 0)
        print(cnpj)
        
        # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
        if usuario_anterior != usuario:
            # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
            if s:
                s.close()
                
            # abre uma nova sessão
            with Session() as s:
                
                erro = 'S'
                contador = 0
                # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
                while erro == 'S':
                    if event == '-encerrar-' or event == sg.WIN_CLOSED:
                        alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                        try:
                            s.close()
                        except:
                            pass
                        try:
                            os.rmdir(pasta_final)
                        except:
                            pass
                        return False
                    if contador >= 3:
                        cookies = 'erro'
                        sid = 'Erro ao logar na empresa'
                        break
                    try:
                        cookies, sid = new_session_fazenda_driver(window_principal, usuario, senha, perfil)
                        erro = 'N'
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                        erro = 'S'
                        contador += 1
                
                if event == '-encerrar-' or event == sg.WIN_CLOSED:
                    alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                    try:
                        s.close()
                    except:
                        pass
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                time.sleep(1)
                # se não salvar os cookies fecha a sessão e vai para o próximo dado
                if cookies == 'erro' or not cookies:
                    texto = f'{cnpj};{sid}'
                    usuario_anterior = 'padrão'
                    s.close()

                    escreve_relatorio_csv(f'{cnpj};{nome};{cidade};{texto}', nome=andamentos, local=pasta_final)

                    # print(f'❗ {sid}\n', end='')
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
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    s.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False
            
            
        # escreve na planilha de andamentos o resultado da execução atual
        texto = f'{cnpj};{str(situacao[2:])}'
        try:
            texto = texto.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')
            escreve_relatorio_csv(f'{cnpj};{nome};{cidade};{texto}', nome=andamentos, local=pasta_final)
        except:
            raise Exception(f"Erro ao escrever esse texto: {texto}")
        print(situacao)
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            try:
                s.close()
            except:
                pass
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return False
        
    return True
    

def run_debitos_divida_ativa(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
    def verifica_debitos(pagina):
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # debito = soup.find('span', attrs={'id': 'consultaDebitoForm:dataTable:tb'})
        soup = soup.prettify()
        # print(soup)
        try:
            request_verification = re.compile(r'(consultaDebitoForm:dataTable:tb)').search(soup).group(1)
            return request_verification
        except:
            re.compile(r'(Nenhum resultado com os critérios de consulta)').search(soup).group(1)
            return False
    
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
        with open('Assets\style.css', 'r') as f:
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
            pisa.CreatePDF(str(html), pdf)
        print('✔ Arquivo salvo')
        escreve_relatorio_csv(f'{cnpj};{empresa};Empresa com debitos;{compl}', nome=andamentos, local=pasta_final)
        
        return True
    
    def inicia_sessao():
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--window-size=1920,1080')
        
        status, driver = initialize_chrome(options)
        
        url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
        driver.get(url)
        
        # gera o token para passar pelo captcha
        recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
        token = solve_recaptcha(recaptcha_data)
        
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
        
        cnpj, empresa = empresa
        print(cnpj)
        
        url = 'https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf'
        # str_cnpj = f"{cnpj[0:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
        pagina = s.get(url)
        soup = BeautifulSoup(pagina.content, 'html.parser')
        # tenta capturar o viewstate do site, pois ele é necessário para a requisição e é um dado que muda a cada acesso
        try:
            viewstate = soup.find(id="javax.faces.ViewState").get('value')
        except Exception as e:
            print('❌ Não encontrou viewState')
            escreve_relatorio_csv(f'{cnpj};{empresa};{e}', nome=andamentos, local=pasta_final)
            print(e)
            print(soup)
            s.close()
            return False
        
        # gera o token para passar pelo captcha
        recaptcha_data = {'sitekey': '6Le9EjMUAAAAAPKi-JVCzXgY_ePjRV9FFVLmWKB_', 'url': url}
        token = solve_recaptcha(recaptcha_data)
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            try:
                s.close()
            except:
                pass
            return
        
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
            'consultaDebitoForm:decLblTipoConsulta:j_id76': 'consultaDebitoForm:decLblTipoConsulta:j_id76'
        }
        s.post(url, info)
        
        # Consulta o cnpj
        info = {
            'consultaDebitoForm': 'consultaDebitoForm',
            'consultaDebitoForm:decLblTipoConsulta:opcoesPesquisa': 'CNPJ',
            'consultaDebitoForm:decTxtTipoConsulta:cnpj': cnpj,
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
            escreve_relatorio_csv(f'{cnpj};{empresa};Empresa sem debitos', nome=andamentos, local=pasta_final)
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
                window_principal['-Mensagens_2-'].update(f'Débitos encontrados, gerando documentos ({index+1} de {len(linhas)})...')
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
    
    colunas_usadas = ['CNPJ', 'Razao']
    colunas_filtro = False
    palavras_filtro = False
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, colunas_filtro=colunas_filtro, palavras_filtro=palavras_filtro)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    window_principal['-Mensagens-'].update('Iniciando ambiente da consulta, aguarde...')
    window_principal.refresh()
    
    # inicia uma sessão com webdriver para gerar cookies no site e garantir que as requisições funcionem depois
    driver, nome_botao_consulta = inicia_sessao()
    # armazena os cookies gerados pelo webdriver
    cookies = driver.get_cookies()
    driver.quit()
    
    # inicia uma sessão para as requisições
    with Session() as s:
        # pega a lista de cookies armazenados e os configura na sessão
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])
        
        tempos = [datetime.now()]
        tempo_execucao = []
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            # printa o indice da empresa que está sendo executada
            tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
            
            # enquanto a consulta não conseguir ser realizada e tenta de novo
            """while True:
            try:"""
            s = consulta_debito(window_principal, s, nome_botao_consulta, empresa, andamentos, pasta_final)
            """break
            except:
                pass"""
            
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    s.close()
                except:
                    pass
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False
    
    # fecha a sessão
    try:
        s.close()
    except:
        pass
    
    return True
    

def run_pendencias_sigiss(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
    
    colunas_usadas = ['CNPJ', 'Senha Prefeitura', 'Razao']
    colunas_filtro = ['Cidade']
    palavras_filtro = ['Valinhos']
    pasta_final, index, empresas = configura_dados(window_principal, codigo_20000, planilha_dados, pasta_final_, andamentos,
                                                   colunas_usadas=colunas_usadas, colunas_filtro=colunas_filtro, palavras_filtro=palavras_filtro)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
            
        cnpj, senha, nome = empresa
        print(cnpj)
        
        nome = nome.replace('/', ' ')
        
        # faz login no SIGISSWEB
        execucao, documento, s = login_sigiss(cnpj, senha)
        if execucao:
            # se fizer login, consulta a situação da empresa
            execucao, documento, s = consulta(s, cnpj, nome, pasta_final)
        
        # escreve os resultados da consulta
        escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}', nome=andamentos, local=pasta_final)
        s.close()
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            alert(f'❗ Rotina "{andamentos}" encerrada pelo usuário')
            try:
                s.close()
            except:
                pass
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return False
        
    return True
        

# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ ROTINAS ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑


def run(window_principal, codigo_20000, planilha_dados, pasta_final, rotina):
    rotinas = {'Consulta Certidão Negativa de Débitos Tributários Não Inscritos':run_cndtni,
               'Consulta Débitos Municipais Jundiaí':run_debitos_municipais_jundiai,
               'Consulta Débitos Municipais Valinhos':run_debitos_municipais_valinhos,
               'Consulta Débitos Municipais Vinhedo':run_debitos_municipais_vinhedo,
               'Consulta Débitos Estaduais - Situação do Contribuinte':run_debitos_estaduais,
               'Consulta Divida Ativa':run_debitos_divida_ativa,
               'Consulta Pendências SIGISSWEB Valinhos':run_pendencias_sigiss}
    
    if rotinas[str(rotina)](window_principal, codigo_20000, planilha_dados, pasta_final, rotina):
        alert(f'✔ Rotina "{rotina}" finalizada')

     
# Define o ícone global da aplicação
sg.set_global_icon(icone)
if __name__ == '__main__':
    tempo_total_segundos = 0
    sg.LOOK_AND_FEEL_TABLE['tema_claro'] = {'BACKGROUND': '#ffffff',
                                      'TEXT': '#0e0e0e',
                                      'INPUT': '#ffffff',
                                      'TEXT_INPUT': '#0e0e0e',
                                      'SCROLL': '#ffffff',
                                      'BUTTON': ('#0e0e0e', '#ffffff'),
                                      'PROGRESS': ('#fca400', '#D7D7D7'),
                                      'BORDER': 0,
                                      'SLIDER_DEPTH': 0,
                                      'PROGRESS_DEPTH': 0}
    
    sg.LOOK_AND_FEEL_TABLE['tema_escuro'] = {'BACKGROUND': '#0e0e0e',
                                            'TEXT': '#ffffff',
                                            'INPUT': '#0e0e0e',
                                            'TEXT_INPUT': '#ffffff',
                                            'SCROLL': '#0e0e0e',
                                            'BUTTON': ('#ffffff', '#0e0e0e'),
                                            'PROGRESS': ('#fca400', '#2A2A2A'),
                                            'BORDER': 0,
                                            'SLIDER_DEPTH': 0,
                                            'PROGRESS_DEPTH': 0}
    
    f = open(dados_modo, 'r', encoding='utf-8')
    modo = f.read()
    sg.theme(f'tema_{modo}')  # Define o tema do PySimpleGUI
    
    def janela_principal():
        def cria_layout():
            rotinas = ['Consulta Certidão Negativa de Débitos Tributários Não Inscritos', 'Consulta Débitos Municipais Jundiaí', 'Consulta Débitos Municipais Valinhos',
                       'Consulta Débitos Municipais Vinhedo', 'Consulta Débitos Estaduais - Situação do Contribuinte', 'Consulta Divida Ativa', 'Consulta Pendências SIGISSWEB Valinhos',]
            
            return [[sg.Button('AJUDA', font=("Helvetica", 10, "underline"), border_width=0),
                 sg.Button('SOBRE', font=("Helvetica", 10, "underline"), border_width=0),
                 sg.Button('LOG DO SISTEMA', font=("Helvetica", 10, "underline"), key='-log_sistema-', border_width=0, disabled=True),
                 sg.Button('CONFIGURAÇÕES', font=("Helvetica", 10, "underline"), key='-config-', border_width=0),
                 sg.Text('', expand_x=True),
                 sg.Button('☀', key='-tema_claro-', border_width=0),
                 sg.Button('☀', key='-tema_escuro-', border_width=0)
                 ],
                [sg.Text('')],
                [sg.Text('')],
                
                [sg.Text('Selecione a rotina que será executada')],
                [sg.Combo(rotinas, expand_x=True, enable_events=True, readonly=True, key='-dropdown-')],
                [sg.Text('')],
                
                [sg.Text('Selecione a pasta para salvar os resultados')],
                [sg.FolderBrowse('SELECIONAR', font=("Helvetica", 10, "underline"), key='-pasta_resultados-', target='-output_dir-'),
                 sg.Text(key='-output_dir-', expand_x=True)],
                [sg.Text('')],
                
                [sg.Text('Selecione a planilha de dados')],
                [sg.FileBrowse('SELECIONAR', font=("Helvetica", 10, "underline"), key='-planilha_dados-', file_types=(('Planilhas Excel', '*.xlsx'),), target='-input_dados-'),
                 sg.Text(key='-input_dados-', expand_x=True)],
                [sg.Text('')],
                
                [sg.Text('Utilizar empresas com o código acima de 20.000')],
                [sg.Checkbox(key='-codigo_20000_sim-', text='Sim', enable_events=True),
                 sg.Checkbox(key='-codigo_20000_nao-', text='Não', enable_events=True, default=True),
                 sg.Checkbox(key='-codigo_20000-', text='Apenas acima do 20.000', enable_events=True)],
                [sg.Text('', expand_y=True)],
                
                [sg.Text('', key='-Mensagens_2-')],
                [sg.Text('', key='-Mensagens-')],
                [sg.Text(size=6, text='', key='-Progresso_texto-'),
                 sg.ProgressBar(max_value=0, orientation='h', size=(5, 5), key='-progressbar-', expand_x=True, visible=False)],
                [sg.Button('INICIAR', font=("Helvetica", 10, "underline"), key='-iniciar-', border_width=0),
                 sg.Button('ENCERRAR', font=("Helvetica", 10, "underline"), key='-encerrar-', disabled=True, border_width=0),
                 sg.Button('ABRIR RESULTADOS', font=("Helvetica", 10, "underline"), disabled=True, key='-abrir_resultados-', border_width=0)],
            ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Consultas ADM', cria_layout(), resizable=True, finalize=True, margins=(30, 30))
    
    def janela_configura():  # layout da janela do menu principal
        def cria_layout_configura():
            return [
                [sg.Text('Insira a nova chave de acesso para a API Anticaptcha')],
                [sg.InputText(key='-input_chave_api-', size=90, password_char='*', default_text='', border_width=1)],
                [sg.Button('APLICAR', key='-confirma_conf-', border_width=0),
                 sg.Button('CANCELAR', key='-cancela_conf-', border_width=0), ]
            ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Configurações', cria_layout_configura(), finalize=True, modal=True, margins=(30, 30))
    
    
    def run_script_thread():
        for elemento in [(rotina, 'Por favor informe uma rotina para executar'), (pasta_final, 'Por favor informe um diretório para salvar os andamentos.'),
                         (planilha_dados, 'Por favor informe uma rotina para executar')]:
            if not elemento[0]:
                alert(f'❗ {elemento[1]}')
                return

        # habilita e desabilita os botões conforme necessário
        for key in [('-tema_claro-', True), ('-tema_escuro-', True), ('-dropdown-', True), ('-planilha_dados-', True), ('-pasta_resultados-', True), ('-codigo_20000_sim-', True),
                    ('-codigo_20000_nao-', True), ('-codigo_20000-', True), ('-iniciar-', True), ('-encerrar-', False), ('-abrir_resultados-', False), ('-config-', True)]:
            window_principal[key[0]].update(disabled=key[1])
            
        # apaga qualquer mensagem na interface
        window_principal['-Mensagens-'].update('')
        window_principal['-Mensagens_2-'].update('')
        # atualiza a barra de progresso para ela ficar mais visível
        window_principal['-progressbar-'].update(visible=True)
        
        try:
            # Chama a função que executa o script
            run(window_principal, codigo_20000, planilha_dados, pasta_final, rotina)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            escreve_doc(f'Traceback: {traceback_str}\n\n'
                        f'Erro: {erro}')
            window_principal['-log_sistema-'].update(disabled=False)
            alert('❌ Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
        
        window_principal['-progressbar-'].update(visible=False)
        window_principal['-progressbar-'].update_bar(0)
        window_principal['-Progresso_texto-'].update('')
        window_principal['-Mensagens-'].update('')
        window_principal['-Mensagens_2-'].update('')
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-tema_claro-', False), ('-tema_escuro-', False), ('-dropdown-', False), ('-planilha_dados-', False), ('-pasta_resultados-', False), ('-codigo_20000_sim-', False), ('-codigo_20000_nao-', False),
                    ('-codigo_20000-', False), ('-iniciar-', False), ('-encerrar-', True), ('-config-', False)]:
            window_principal[key[0]].update(disabled=key[1])
        
    # inicia as variáveis das janelas
    window_principal, window_configura = janela_principal(), None
    # Definindo o tamanho mínimo da janela
    window_principal.set_min_size((600, 600))
    
    codigo_20000 = '-codigo_20000_nao-'
    while True:
        # captura o evento e os valores armazenados na interface
        window, event, values = sg.read_all_windows()
        
        checkboxes = ['-codigo_20000_sim-', '-codigo_20000_nao-', '-codigo_20000-']
        
        try:
            pasta_final = values['-pasta_resultados-']
            rotina = values['-dropdown-']
            planilha_dados = values['-planilha_dados-']
        except:
            pasta_final = None
            rotina = None
            planilha_dados = None
        
        if event in ('-codigo_20000_sim-', '-codigo_20000_nao-', '-codigo_20000-'):
            for checkbox in checkboxes:
                if checkbox != event:
                    window[checkbox].update(value=False)
                else:
                    codigo_20000 = checkbox
        
        if event == sg.WIN_CLOSED:
            if window == window_configura:  # if closing win 2, mark as closed
                window_configura = None
            elif window == window_principal:  # if closing win 1, exit program
                break
        
        elif event == '-tema_claro-' or event == '-tema_escuro-':
            nome_tema = event.replace('-', '')
            controle_tema = nome_tema.split('_')[1]
            sg.theme(nome_tema)  # Define o tema claro
            with open(dados_modo, 'w', encoding='utf-8') as f:
                f.write(controle_tema)
            window.close()  # Fecha a janela atual
            # inicia as variáveis das janelas
            window_principal, window_configura = janela_principal(), None
            # Definindo o tamanho mínimo da janela
            window_principal.set_min_size((600, 600))
        
        elif event == '-config-':
            window_configura = janela_configura()
            
            while True:
                # captura o evento e os valores armazenados na interface
                event, values = window_configura.read()
                if event == '-confirma_conf-':
                    confirma = confirm(text='As alterações serão aplicadas, deseja continuar?')
                    nova_chave = values['-input_chave_api-']
                    if confirma == 'OK':
                        with open(dados_anticaptcha, 'w', encoding='utf-8') as f:
                            f.write(nova_chave)
                        break
                if event == '-cancela_conf-':
                    confirma = confirm(text='As alterações serão perdidas, deseja continuar?')
                    if confirma == 'OK':
                        break
                if event == sg.WIN_CLOSED:
                    confirma = confirm(text='As alterações serão perdidas, deseja continuar?')
                    if confirma == 'OK':
                        break
                    window_configura = janela_configura()
            
            window_configura.close()
            window_principal.UnHide()
        
        elif event == '-log_sistema-':
            os.startfile('Log')
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Cria E-book DIRPF.pdf')
        
        elif event == 'Sobre':
            os.startfile('Sobre.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
            
        elif event == '-encerrar-':
            window_principal['-Mensagens_2-'].update('')
            window_principal['-Mensagens-'].update('Encerrando, aguarde...')
            
        elif event == '-abrir_resultados-':
            os.makedirs(os.path.join(pasta_final, rotina), exist_ok=True)
            os.startfile(os.path.join(pasta_final, rotina))
        
    window_principal.close()
