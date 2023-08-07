# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pyperclip, os, time, re, csv, shutil, fitz, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome, _find_by_path, _find_by_id
from pyautogui_comum import _click_img, _find_img, _wait_img


def login_sieg(driver):
    driver.get('https://auth.sieg.com/')
    print('>>> Acessando o site')
    time.sleep(1)

    dados = "V:\\Setor Robô\\Scripts Python\\_comum\\DadosVeri-SIEG.txt"
    f = open(dados, 'r', encoding='utf-8')
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
    
    return driver
    

def sieg_iris(driver, modulo):
    print(f'>>> Acessando IriS {modulo}')
    if modulo == 'Termo(s) de Exclusão do Simples Nacional':
        driver.get('https://hub.sieg.com/IriS/#/Mensagens?Filtro=Exclusoes')
    if modulo == 'Termo(s) de Intimação':
        driver.get('https://hub.sieg.com/IriS/#/Mensagens?Filtro=Intimacoes')
    
    return driver


def imprime_mensagem(driver):
    print('>>> Acessando mensagem')
    # aguarda a lista de mensagens
    while not _find_by_path('/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]', driver):
        if _find_img('sem_mensagens.png', conf=0.8):
            return False
        time.sleep(1)
        
    # aguarda e clica para filtrar abrir o dropdown de filtro de mensagens não lidas
    while not _find_by_id('select2-ddlSelectedMessage-container', driver):
        time.sleep(1)
    
    while not _find_img('filtros_carregados.png', conf=0.8):
        time.sleep(1)
        
    time.sleep(1)
    _click_img('filtro_lidas.png', conf=0.9)
    _wait_img('opcao_filtro_n_lidas.png', conf=0.9)
    _click_img('opcao_filtro_n_lidas.png', conf=0.9)
    time.sleep(1)
    
    # aguarda e clica na mensagem
    while not _find_by_path('/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]', driver):
        if _find_img('sem_mensagens.png', conf=0.8):
            return False
        time.sleep(1)
    
    return driver
    

def salvar_pdf(driver, pasta_analise):
    time.sleep(1)
    # enquanto a pre visualização nao abrir tenta abrir a primeira mensagem da lista
    print('>>> Aguardando pré visualização')
    while not _find_img('pre_visualizacao.png', conf=0.75):
        while True:
            try:
                if _find_img('sem_mensagens.png', conf=0.8):
                    return driver, 'acabou'
                
                driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div/div[1]/div[3]/div[2]/div/table/tbody/tr[1]').click()
                
                break
            except:
                pass
        
        time.sleep(2)
    
    # aguarda a janela de pré-visualização abrir e clica em imprimir
    while not _find_img('imprimir.png', conf=0.8):
        time.sleep(1)
        _click_img('pre_visualizacao.png', conf=0.8)
        time.sleep(1)
        p.press('pgdn')
        time.sleep(1)
        
    _click_img('imprimir.png', conf=0.8, timeout=1)

    print('>>> Aguardando tela de impressão')
    # aguarda a tela de impressão
    _wait_img('tela_imprimir.png', conf=0.9)
    
    print('>>> Salvando PDF')
    # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
    imagens = ['print_to_pdf.png', 'print_to_pdf_2.png']
    for img in imagens:
        if _find_img(img, conf=0.9) or _find_img(img, conf=0.9):
            _click_img(img, conf=0.9)
            # aguarda aparecer a opção de salvar como PDF e clica nela
            _wait_img('salvar_como_pdf.png', conf=0.9)
            _click_img('salvar_como_pdf.png', conf=0.9)
    
    # aguarda aparecer o botão de salvar e clica nele
    _wait_img('botao_salvar.png', conf=0.9)
    _click_img('botao_salvar.png', conf=0.9)
    
    print('>>> Salvando relatório')
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
    
    os.makedirs(pasta_analise, exist_ok=True)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    while True:
        try:
            pyperclip.copy(pasta_analise)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(1)
    
    print('✔ Mensagem salva com sucesso')
    return driver, 'Mensagem salva com sucesso'


def verifica_mensagem(pasta_analise, pasta_final, modulo):
    print('>>> Analisando mensagem')
    # Analisa cada pdf que estiver na pasta
    for arquivo in os.listdir(pasta_analise):
        # Abrir o pdf
        arq = os.path.join(pasta_analise, arquivo)
        
        with fitz.open(arq) as pdf:
            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                # se a função for chamada com o parametro 'erro', printa o texto do arquivo para ser analisado e retorna para encerrar a execução
                if modulo == 'erro':
                    print(textinho)
                    return
                
                try:
                    cnpj = re.compile(r'Destinatário: (.+)').search(textinho).group(1)
                except:
                    cnpj = re.compile(r'CPF (\d\d\d\.\d\d\d\.\d\d\d-\d\d) ').search(textinho).group(1)
                    cnpj = cnpj.replace("-", "").replace(".", "").replace("/", "")
                    
                # verifica código e data da mensagem para caso tenha a mesma empresa o arquivo anterior não seja substituído
                regexes = [r'SIMPLES NACIONAL.+nº (.+), de (.+)\n',
                           r'SIMPLES NACIONAL.+Nº (.+), DE (.+)\.',
                           r'Número da Intimação: (\d+).+\n(\d\d/\d\d/\d\d\d\d)',
                           r'PER/DCOMP (.+) - ',
                           r'nº (.+) DE (.+)',
                           r'nº (.+) de (.+)']
                for regex in regexes:
                    info_termo = re.compile(regex).search(textinho)
                    
                    if info_termo:
                        numero = info_termo.group(1)
                        numero = numero.replace("/", "-")
                        
                        # caso existe uma variação onde o regex não consiga pegar o código e a data juntos entra no except para pegar somente a data
                        try:
                            data = info_termo.group(2)
                        except:
                            data = re.compile(r'Data de envio: (\d\d/\d\d/\d\d\d\d) ').search(textinho).group(1)
                        data = data.replace("/", "-").replace(".", "")
                        
                        novo_arq = f'{cnpj} - {modulo} - {numero} - {data}.pdf'
                        
                        print(f"{novo_arq.replace('.pdf', '')}\n")
                        break
                        
                # pega o conteúdo mais relevante do termo de intimação para análise futura
                
                if modulo == 'Termo(s) de Intimação':
                    regexes = [r'Data da Emissão: \d\d/\d\d/\d\d\d\d(.+)',
                               r'Pela presente mensagem, (.+)',
                               r'(Fica o sujeito passivo.+)',]
                    for regex in regexes:
                        info_termo = re.compile(regex).search(textinho.replace('\n', ' '))
                        
                        if info_termo:
                            mensagem = info_termo.group(1)
                            mensagem = (mensagem.replace(' ; ', ' - ')
                                           .replace('; ', ' ')
                                           .replace('arresto;bens', 'arresto de bens')
                                           .replace(';', ' ')
                                           .replace(';;', ' '))
                            break
                else:
                    mensagem = ''
                    
                _escreve_relatorio_csv(novo_arq.replace(' - ', ';').replace('.pdf', '') + ';' + mensagem)
                break
                        
        # verifica se já existe a pasta final do arquivo e move ele para lá
        os.makedirs(pasta_final, exist_ok=True)
        shutil.move(arq, os.path.join(pasta_final, novo_arq))


@_time_execution
def run():
    modulo = p.confirm(title='Script incrível', buttons=('Termo(s) de Exclusão do Simples Nacional', 'Termo(s) de Intimação', 'Renomear arquivo'))
    pasta_analise = r'V:\Setor Robô\Scripts Python\SIEG\Download Menságens Importantes\ignore\Mensagens'
    pasta_final = os.path.join(r'V:\Setor Robô\Scripts Python\SIEG\Download Menságens Importantes\execução', modulo)
    
    # exceção para renomear alguma menagem que foi salva ou só aberta, mas o script parou
    if modulo == 'Renomear arquivo':
        modulo = p.confirm(title='Script incrível', buttons=('Termo(s) de Exclusão do Simples Nacional', 'Termo(s) de Intimação'))
        pasta_final = os.path.join(r'V:\Setor Robô\Scripts Python\SIEG\Download Menságens Importantes\execução', modulo)
        verifica_mensagem(pasta_analise, pasta_final, modulo)
        return
        
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.9")
    
    # iniciar o driver do chrome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)
    
    while True:
        driver = sieg_iris(driver, modulo)
        
        if imprime_mensagem(driver):
            driver, resultado = salvar_pdf(driver, pasta_analise)
            if resultado == 'acabou':
                driver.close()
                return
                
            # tenta analisar o PDF, se der erro chama a função novamente, porem com o parametro 'erro' para entrar em uma exceção criada dentro da função
            try:
                verifica_mensagem(pasta_analise, pasta_final, modulo)
            except:
                verifica_mensagem(pasta_analise, 'erro')
                return
                
            
            
            
if __name__ == '__main__':
    run()
