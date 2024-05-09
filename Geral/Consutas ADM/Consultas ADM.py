# -*- coding: utf-8 -*-
import chromedriver_autoinstaller, datetime, shutil, time, io, re, os, sys, traceback, pandas as pd, PySimpleGUI as sg
from win32com import client
from requests import Session
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from twocaptcha import TwoCaptcha
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from PIL import Image
from threading import Thread
from pyautogui import confirm
from PyQt5.QtWidgets import QApplication, QMessageBox, QLabel
from PyQt5.QtGui import QPixmap, QIcon

icone = 'Assets/auto-flash.ico'
dados_anticaptcha = "Dados\Dados 2Cap.txt"


def exibir_alerta(imagem, texto):
    imagem = os.path.join('Assets', imagem)
    # Crie a aplicação
    app = QApplication([])
    
    # Crie uma caixa de mensagem
    msg = QMessageBox()
    msg.setWindowTitle(" ")
    msg.setText(texto)
    msg.setWindowIcon(QIcon('Assets/em_branco.ico'))
    
    # Adicione uma imagem à mensagem
    pixmap = QPixmap(imagem)
    msg.setIconPixmap(pixmap)
    
    # Mostre a caixa de mensagem
    msg.exec_()



def concatena(variavel, quantidade, posicao, caractere):
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
    chromedriver_autoinstaller.install()
    print('>>> Inicializando Chromedriver...')
    
    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
    
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
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


def solve_recaptcha(data):
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_2cap = f.read()
    api_key = os.getenv('APIKEY_2CAPTCHA', key_2cap)
    tipo = 'recaptcha-v2'
    print(f'>>> Resolvendo {tipo}')
    
    solver = TwoCaptcha(api_key)
    result = solver.recaptcha(sitekey=data['sitekey'], url=data['url'])
    
    print('>>> Captcha resolvido')
    try:
        return str(result['code'])
    except:
        print(result)


def solve_text_captcha(driver, captcha_element, element_type='id'):
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_2cap = f.read()
    api_key = os.getenv('APIKEY_2CAPTCHA', key_2cap)
    tipo = 'text captcha'
    print(f'>>> Resolvendo {tipo}')
    
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
    
    print('>>> Resolvendo text captcha')
    
    solver = TwoCaptcha(api_key)
    result = solver.normal(os.path.join('Log', 'captcha', 'captcha.png'))
    
    print('>>> Captcha resolvido')
    try:
        return str(result['code'])
    except:
        print(result)


def indice(count, empresa, total_empresas, index, window_princial, tempos, tempo_execucao):
    tempo_estimado = 0
    tempo_inicial = datetime.datetime.now()
    
    tempos.append(tempo_inicial)
    tempo_execucao_atual = int(tempos[1].timestamp()) - int(tempos[0].timestamp())
    
    tempo_execucao.append(tempo_execucao_atual)
    for t in tempo_execucao:
        tempo_estimado = tempo_estimado + t
    tempo_estimado = int(tempo_estimado) / int(len(tempo_execucao))
    
    tempo_total_segundos = int((len(total_empresas) + index) - (count + index) + 1) * float(tempo_estimado)
    
    # Converter o tempo total para um objeto timedelta
    tempo_total = datetime.timedelta(seconds=tempo_total_segundos)
    
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
    window_princial['-Mensagens-'].update(f'{str((count + index) - 1)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index) + 1)} Restantes{tempo_estimado_texto}')
    window_princial['-progressbar-'].update_bar(count, max=int(len(total_empresas)))
    window_princial['-Progresso_texto-'].update(str(round(float(count) / int(len(total_empresas)) * 100, 1)) + '%')
    window_princial.refresh()
    
    tempo_estimado = tempo_execucao
    return tempos, tempo_estimado


# wrapper para askopenfilename
def ask_for_file(title='Abrir arquivo', filetypes='*', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    return file if file else False


def where_to_start(idents, pasta_final_anterior, planilha, encode='latin-1'):
    if not os.path.isdir(pasta_final_anterior):
        return 0
        
    file = os.path.join(pasta_final_anterior, planilha)

    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except:
        # alert(title='❕', text=f'Não foi encontrada nenhuma planilha de andamentos na pasta de execução anterior.\n\n'
        #            f'Começando a execução a partir do primeiro indice da planilha de dados selecionada.')
        exibir_alerta('i.png', f'Não foi encontrada nenhuma planilha de andamentos na pasta de execução anterior.\n\n'
                   f'Começando a execução a partir do primeiro indice da planilha de dados selecionada.')
        return 0
    
    try:
        elem = dados[-1].split(';')[0]
        if len(idents) == idents.index(elem) + 1:
            return 0
        return idents.index(elem) + 1
    except ValueError:
        return 0


def open_dados(codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro):
    def open_lista_dados(df_filtrada, encode='latin-1'):
        try:
            os.remove('Dados.csv')
        except:
            pass
        
        df_filtrada.to_csv('Dados.csv', header=False, index=False, sep=';')
        try:
            with open('Dados.csv', 'r', encoding=encode) as f:
                dados = f.readlines()
        except Exception as e:
            # alert(title='❌', text=f'Não pode abrir arquivo\n{planilha_dados}\n{str(e)}')
            exibir_alerta('x.png', f'Não pode abrir arquivo\n{planilha_dados}\n{str(e)}')
            return False
        
        return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))
    
    # modelo de lista com as colunas que serão usadas na rotina
    # colunas_usadas = ['column1', 'column2', 'column3']
    
    df = pd.read_excel(planilha_dados)
    
    # coluna com os códigos do ae
    coluna_codigo = 'Codigo'
    
    if codigo_20000 == '-codigo_20000_nao-':
        # cria um novo df apenas com empresas a baixo do código 20.000
        df_filtrada = df[df[coluna_codigo] <= 20000]
    elif codigo_20000 == '-codigo_20000-':
        # cria um novo df apenas com empresas a cima do código 20.000
        df_filtrada = df[df[coluna_codigo] >= 20000]
    else:
        df_filtrada = df
        
    # filtra as células de colunas específicas que contenham palavras especificas
    if palavras_filtro and colunas_filtro:
        for count, coluna_para_filtrar in enumerate(colunas_filtro):
            df_filtrada = df_filtrada[df_filtrada[coluna_para_filtrar].str.contains(palavras_filtro[count], case=False, na=False)]
    
    # filtra as colunas
    try:
        df_filtrada = df_filtrada[colunas_usadas]
    except KeyError:
        # alert(title='❌', text=f'Erro ao buscar as colunas na planilha base selecionada: {planilha_dados}\n\n'
        #                      f'Verifique se a planilha contem as colunas necessárias para a execução da rotina e se elas tem exatamente o mesmo nome indicado ao lado: {colunas_usadas}')
        exibir_alerta('x.png', f'Erro ao buscar as colunas na planilha base selecionada: {planilha_dados}\n\n'
                              f'Verifique se a planilha contem as colunas necessárias para a execução da rotina e se elas tem exatamente o mesmo nome indicado ao lado: {colunas_usadas}')
        return False
        
    # remove linha com células vazias
    df_filtrada = df_filtrada.dropna(axis=0, how='any')

    # Converta a coluna 'codigo' para string. Converta a coluna 'codigo' para string e remova a parte decimal '.0'. Preencha com zeros à esquerda para garantir 14 dígitos
    df_filtrada['CNPJ'] = df_filtrada['CNPJ'].astype(str).str.replace(r'\.0', '', regex=True).str.zfill(14)

    empresas = open_lista_dados(df_filtrada, encode='latin-1')
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
    
    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        try:
            f = open(os.path.join(local, f"{nome} - Parte 2.csv"), 'a', encoding=encode)
        except:
            f = open(os.path.join(local, f"{nome} - Parte 3.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()
    
    
def configura_dados(window_princial, codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, pasta_final_, andamentos):
    comp = datetime.datetime.now().strftime('%m-%Y')
    pasta_final_ = os.path.join(pasta_final_, andamentos, comp)
    contador = 0
    while True:
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
    
    window_princial['-Mensagens-'].update('Criando dados para a consulta...')
    window_princial.refresh()
    # abrir a planilha de dados
    empresas = open_dados(codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro)
    if not empresas:
        return '', '', False
    
    if pasta_final_anterior:
        planilha = f'{andamentos}.csv'
        index = where_to_start(tuple(i[0] for i in empresas), pasta_final_anterior, planilha)
    else:
        index = 0
    
    return pasta_final, index, empresas


# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ FUNÇÕES PADRÕES ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ ROTINAS ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓


def run_debitos_municipais_valinhos(window_princial, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
    pasta_final, index, empresas = configura_dados(window_princial, codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, pasta_final_, andamentos)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_princial, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome = empresa
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        try:
            insc_muni = int(insc_muni)
        except:
            continue
        
        resultado = 'Texto da imagem incorreto'
        while resultado == 'Texto da imagem incorreto':
            window_princial['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_princial.refresh()
            
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
            # alert(title='❕', text=f'Rotina "{andamentos}" encerrada pelo usuário')
            exibir_alerta('i.png', f'Rotina "{andamentos}" encerrada pelo usuário')
            try:
                driver.close()
            except:
                pass
            
            try:
                os.rmdir(pasta_final)
            except:
                pass
            return False
    
    return True
    
   
def run_debitos_municipais_vinhedo(window_princial, codigo_20000, planilha_dados, pasta_final_, andamentos):
    def pesquisar_vinhedo(window_princial, options, cnpj, insc_muni, pasta_final):
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
        window_princial['-Mensagens_2-'].update('Consultando empresa...')
        window_princial.refresh()
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
        
        situacao = salvar_guia_vinhedo(window_princial, driver, cnpj, pasta_final)
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
    
    def salvar_guia_vinhedo(window_princial, driver, cnpj, pasta_final):
        window_princial['-Mensagens_2-'].update('Buscando Certidão Negativa...')
        window_princial.refresh()
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
        
        window_princial['-Mensagens_2-'].update('Salvando Certidão Negativa...')
        window_princial.refresh()
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
    pasta_final, index, empresas = configura_dados(window_princial, codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, pasta_final_, andamentos)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    print(pasta_final.replace('/', '\\'))
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
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_princial, tempos, tempo_execucao)

        cnpj, insc_muni, nome = empresa
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        while True:
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                # alert(title='❕', text=f'Rotina "{andamentos}" encerrada pelo usuário')
                exibir_alerta('i.png', f'Rotina "{andamentos}" encerrada pelo usuário')
                try:
                    os.rmdir(pasta_final)
                except:
                    pass
                return False
            
            window_princial['-Mensagens_2-'].update('Entrando no perfil da empresa...')
            window_princial.refresh()
            
            situacao, situacao_print = pesquisar_vinhedo(window_princial, options, cnpj, insc_muni, pasta_final)
            if situacao_print != 'recomeçar':
                print(situacao_print)
                if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
                    # alert(title='❗', text=f'Rotina "{andamentos}":\n\nDesculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.')
                    exibir_alerta('i.png', f'Rotina "{andamentos}":\n\nDesculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.')
                    try:
                        os.rmdir(pasta_final)
                    except:
                        pass
                    return False
                
                escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{situacao}', nome=andamentos, local=pasta_final)
                break
    
    return True


def run_debitos_divida_ativa(window_princial, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
    
    def consulta_debito(window_princial, s, nome_botao_consulta, empresa, andamentos, pasta_final):
        window_princial['-Mensagens_2-'].update('Entrando no perfil da empresa...')
        window_princial.refresh()
        
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
        window_princial['-Mensagens_2-'].update('Verificando débitos...')
        window_princial.refresh()
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
                window_princial['-Mensagens_2-'].update(f'Débitos encontrados, gerando documentos ({index+1} de {len(linhas)})...')
                window_princial.refresh()
                
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
    pasta_final, index, empresas = configura_dados(window_princial, codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, pasta_final_, andamentos)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    window_princial['-Mensagens-'].update('Iniciando ambiente da consulta, aguarde...')
    window_princial.refresh()
    
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
        
        tempos = [datetime.datetime.now()]
        tempo_execucao = []
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            # printa o indice da empresa que está sendo executada
            tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_princial, tempos, tempo_execucao)
            
            # enquanto a consulta não conseguir ser realizada e tenta de novo
            """while True:
            try:"""
            s = consulta_debito(window_princial, s, nome_botao_consulta, empresa, andamentos, pasta_final)
            """break
            except:
                pass"""
            
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                # alert(title='❕', text=f'Rotina "{andamentos}" encerrada pelo usuário')
                exibir_alerta('i.png', f'Rotina "{andamentos}" encerrada pelo usuário')
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
    

def run_pendencias_sigiss(window_princial, codigo_20000, planilha_dados, pasta_final_, andamentos):
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
    pasta_final, index, empresas = configura_dados(window_princial, codigo_20000, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, pasta_final_, andamentos)
    if not empresas:
        try:
            os.rmdir(pasta_final)
        except:
            pass
        return False
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = indice(count, empresa, total_empresas, index, window_princial, tempos, tempo_execucao)
            
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
            # alert(title='❕', text=f'Rotina "{andamentos}" encerrada pelo usuário')
            exibir_alerta('i.png', f'Rotina "{andamentos}" encerrada pelo usuário')
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


def run(window_princial, codigo_20000, planilha_dados, pasta_final, rotina):
    rotinas = {'Consulta Débitos Municipais Valinhos':run_debitos_municipais_valinhos,
               'Consulta Débitos Municipais Vinhedo':run_debitos_municipais_vinhedo,
               'Consulta Divida Ativa':run_debitos_divida_ativa,
               'Consulta Pendências SIGISSWEB Valinhos':run_pendencias_sigiss}
    
    if rotinas[str(rotina)](window_princial, codigo_20000, planilha_dados, pasta_final, rotina):
        # alert(title='✔', text=f'Rotina "{andamentos}" finalizada')
        exibir_alerta('ok.png', f'Rotina "{rotina}" finalizada')

    
    
# Define o ícone global da aplicação
sg.set_global_icon(icone)
if __name__ == '__main__':
    tempo_total_segundos = 0
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    
    def janela_principal():
        # 'Consulta Certidão Negativa de Débitos Tributários Não Inscritos',
        # 'Consulta Débitos Municipais Jundiaí',
        # 'Débitos Estaduais / Situação do Contribuinte'
        rotinas = ['Consulta Débitos Municipais Valinhos', 'Consulta Débitos Municipais Vinhedo', 'Consulta Divida Ativa', 'Consulta Pendências SIGISSWEB Valinhos',]
        
        layout = [
            [sg.Button('Ajuda', border_width=0), sg.Button('Sobre', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True), sg.Button('Configurações', key='-config-', border_width=0)],
            [sg.Text('')],
            [sg.Text('Selecione a rotina que será executada')],
            [sg.Combo(rotinas, expand_x=True, enable_events=True, readonly=True, key='-dropdawn-', text_color='#fca400')],
            [sg.Text('')],
            [sg.Text('Selecione a pasta para salvar os resultados')],
            [sg.FolderBrowse('Pesquisar', key='-pasta_resultados-', target='-output_dir-'), sg.InputText(key='-output_dir-', size=80, disabled=True)],
            [sg.Text('Selecione a planilha de dados')],
            [sg.FileBrowse('Pesquisar', key='-planilha_dados-', file_types=(('Planilhas Excel', '*.xlsx'),), target='-input_dados-'), sg.InputText(key='-input_dados-', size=80, disabled=True)],
            [sg.Text('')],
            [sg.Text('Utilizar empresas com o código acima de 20.000')],
            [sg.Checkbox(key='-codigo_20000_sim-', text='Sim', enable_events=True), sg.Checkbox(key='-codigo_20000_nao-', text='Não', enable_events=True, default=True), sg.Checkbox(key='-codigo_20000-', text='Apenas acima do 20.000', enable_events=True)],
            [sg.Text('')],
            [sg.Text('', key='-Mensagens_2-')],
            [sg.Text('', key='-Mensagens-')],
            [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color=('#fca400', '#ffe0a6'), visible=False)],
            [sg.Button('Iniciar', key='-iniciar-', border_width=0), sg.Button('Encerrar', key='-encerrar-', disabled=True, border_width=0), sg.Button('Abrir resultados', disabled=True, key='-abrir_resultados-', border_width=0)],
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Consultas ADM', layout, finalize=True)
    
    def janela_configura():  # layout da janela do menu principal
        layout_configura = [
            [sg.Text('Insira a nova chave de acesso para a API Anticaptcha', justification='center')],
            [sg.InputText(key='-input_chave_api-', size=90, password_char='*', default_text='')],
            [sg.Button('Aplicar', key='-confirma_conf-', size=30, border_width=1),
             sg.Button('Cancelar', key='-cancela_conf-', size=30, border_width=1), ]
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Configurações', layout_configura, finalize=True, modal=True)
    
    
    def run_script_thread():
        for elemento in [(rotina, 'Por favor informe uma rotina para executar'), (pasta_final, 'Por favor informe um diretório para salvar os andamentos.'),
                         (planilha_dados, 'Por favor informe uma rotina para executar')]:
            if not elemento[0]:
                # alert(title='❌', text=elemento[1])
                exibir_alerta('x.png', elemento[1])
                return

        # habilita e desabilita os botões conforme necessário
        for key in [('-dropdawn-', True), ('-planilha_dados-', True), ('-pasta_resultados-', True), ('-codigo_20000_sim-', True), ('-codigo_20000_nao-', True),
                    ('-codigo_20000-', True), ('-iniciar-', True), ('-encerrar-', False), ('-abrir_resultados-', False), ('-config-', True)]:
            window_princial[key[0]].update(disabled=key[1])
            
        # apaga qualquer mensagem na interface
        window_princial['-Mensagens-'].update('')
        window_princial['-Mensagens_2-'].update('')
        # atualiza a barra de progresso para ela ficar mais visível
        window_princial['-progressbar-'].update(visible=True)
        
        try:
            # Chama a função que executa o script
            run(window_princial, codigo_20000, planilha_dados, pasta_final, rotina)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            escreve_doc(f'Traceback: {traceback_str}\n\n'
                        f'Erro: {erro}')
            window_princial['Log do sistema'].update(disabled=False)
            # alert(title='❌', text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
            exibir_alerta('x.png', 'Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
        
        window_princial['-progressbar-'].update(visible=False)
        window_princial['-progressbar-'].update_bar(0)
        window_princial['-Progresso_texto-'].update('')
        window_princial['-Mensagens-'].update('')
        window_princial['-Mensagens_2-'].update('')
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-dropdawn-', False), ('-planilha_dados-', False), ('-pasta_resultados-', False), ('-codigo_20000_sim-', False), ('-codigo_20000_nao-', False),
                    ('-codigo_20000-', False), ('-iniciar-', False), ('-encerrar-', True), ('-config-', False)]:
            window_princial[key[0]].update(disabled=key[1])
        
    # inicia as variáveis das janelas
    window_princial, window_configura = janela_principal(), None
    
    codigo_20000 = None
    while True:
        # captura o evento e os valores armazenados na interface
        window, event, values = sg.read_all_windows()
        
        checkboxes = ['-codigo_20000_sim-', '-codigo_20000_nao-', '-codigo_20000-']
        
        try:
            pasta_final = values['-pasta_resultados-']
            rotina = values['-dropdawn-']
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
            elif window == window_princial:  # if closing win 1, exit program
                break
        
        elif event == '-config-':
            window_princial.Hide()
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
            window_princial.UnHide()
        
        elif event == 'Log do sistema':
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
            window_princial['-Mensagens_2-'].update('')
            window_princial['-Mensagens-'].update('Encerrando, aguarde...')
            
        elif event == '-abrir_resultados-':
            os.makedirs(os.path.join(pasta_final, rotina), exist_ok=True)
            os.startfile(os.path.join(pasta_final, rotina))
        
    window_princial.close()
