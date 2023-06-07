# -*- coding: utf-8 -*-
import datetime
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from pathlib import Path
import random, time, re, pywhatkit, pandas as pd, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers, _remove_emojis
from chrome_comum import _initialize_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados e-mail.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')

email = user[0]
senha = user[1]


def localiza_path(driver, elemento):
    try:
        driver.find_element(by=By.XPATH, value=elemento)
        return True
    except:
        return False


def localiza_id(driver, elemento):
    try:
        driver.find_element(by=By.ID, value=elemento)
        return True
    except:
        return False


def login_email(driver):
    driver.get('https://smartermail.hiway.com.br/')
    print('>>> Acessando o site')
    time.sleep(2)
    driver.find_element(by=By.ID, value='loginUsernameBox').send_keys(email)
    driver.find_element(by=By.ID, value='loginPasswordBox').send_keys(senha)
    time.sleep(1)
    
    driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/div[1]/div[1]/form/div[2]/div[2]/button').click()
    
    _wait_img('menu_classificacao.png', conf=0.9)
    _click_img('menu_classificacao.png', conf=0.9)
    _wait_img('crescente.png', conf=0.9)
    _click_img('crescente.png', conf=0.9)
    _click_img('menu_classificacao.png', conf=0.9)
    
    return driver


def captura_link_email(driver):
    print('>>> Capturando dados da mensagem')
    
    while not localiza_id(driver, 'messageViewFrame'):
        time.sleep(1)
    
    titulo = re.compile(r'<div id=\"subject\">(.+)</div>').search(driver.page_source).group(1)
    titulo = titulo.replace('-&nbsp;', '- ').replace(' &nbsp;', ' ').replace('&nbsp; ', ' ').replace('&nbsp;', ' ').replace('&amp;', '&')
    
    print(titulo)
    
    # Encontra o elemento do frame pelo nome ou índice
    frame = driver.find_element(by=By.ID, value='messageViewFrame')
    # frame = driver.find_element_by_index(0)
    
    # Alterna o driver para o contexto do frame
    driver.switch_to.frame(frame)
    
    try:
        # pega cnpj da empresa que vai receber a mensagem
        cnpj = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)').search(driver.page_source).group(1)
        cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
        print(cnpj)
        print(cnpj_limpo)
    except:
        try:
            # pega cnpj da empresa que vai receber a mensagem
            cnpj = re.compile(r'CNPJ:</strong> (\d\d\d\d\d\d\d\d\d\d\d)&nbsp;').search(driver.page_source).group(1)
            cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
            print(cnpj)
            print(cnpj_limpo)
        except:
            return driver, titulo, 'x', 'x', 'x', 'x'
        
    # pega a data de vencimento do documento que está no link da mensagem
    try:
        vencimento = re.compile(r'Vencimento:.+(\d\d/\d\d/\d\d\d\d)').search(driver.page_source).group(1)
        vencimento = f"Vencimento: {vencimento}"
    except:
        vencimento = ''
    print(vencimento)
    
    # pega o link que será enviado na mensagem
    link_email_padrao = re.compile(r'<a href=\"(https://api\.gestta\.com\.br.+)\" target=\"_blank\"').findall(driver.page_source)
    if not link_email_padrao:
        link_email_padrao = re.compile(r'<a href=\"(https://app\.gestta\.com\.br.+)\" target=\"_blank\"').findall(driver.page_source)
    
    link_email = list(set(link_email_padrao))
    
    link_mensagem = ''
    contador = 1
    for link in link_email:
        link_mensagem += str(contador) + ' - ' + link + '\n'
        contador += 1
    
    link_mensagem = link_mensagem.replace(' ', '').replace('&amp;', '&')
    print(link_mensagem)
    return driver, titulo, cnpj, cnpj_limpo, vencimento, str(link_mensagem)


def verifica_o_numero(cnpj):
    try:
        print('>>> Consultando número de telefone do cliente')
        # Definir o nome das colunas
        colunas = ['cnpj', 'numero', 'nome']
        # Localiza a planilha
        caminho = Path('V:/Setor Robô/Scripts Python/Geral/Envia Documento/ignore/Dados.csv')
        # Abrir a planilha
        lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
        # Definir o index da planilha
        lista.set_index('cnpj', inplace=True)
        
        cliente = lista.loc[int(cnpj)]
        cliente = list(cliente)
        numero = cliente[0]
        nome = cliente[1]
        
        return str(numero), str(nome)
    except:
        return False, False


def envia(numero, titulo, vencimento, link_mensagem):
    print('>>> Enviando mensagem')
    try:
        pywhatkit.sendwhatmsg_instantly('+55' + numero, f"Olá!\n"
                                                        f"{titulo}\n"
                                                        f"{vencimento}\n"
                                                        f"{link_mensagem}\n"
                                                        f"Obrigado,\n"
                                                        f"R.POSTAL SERVICOS CONTABEIS LTDA\n"
                                                        f"veigaepostal@veigaepostal.com.br\n"
                                                        f"(19)3829-8959", 30, True, 3)
        
        time.sleep(2)
        # verifica se o script secundário localizou algum erro no whatsapp
        arquivos = os.listdir(os.path.join('ignore', 'controle'))
        for arquivo in arquivos:
            resultado = arquivo.split('.')
            resultado = resultado[0]
            os.remove(os.path.join('ignore', 'controle', arquivo))
            return resultado

        return 'ok'
        
    except:
        return 'erro'
    

def mover_email(pasta=''):
    # clica no menu
    _click_img('menu_email.png', conf=0.9)
    
    # espera o menu abrir e clica na opção de mover
    _wait_img('mover_email.png', conf=0.9)
    _click_img('mover_email.png', conf=0.9, clicks=2)
    
    # aguarda a tela para selecionar e depois clica para abrir a lista de caixas
    _wait_img('tela_mover_email.png', conf=0.9)
    _click_img('lista_caixas_email.png', conf=0.9)
    
    # enquanto a caixa não aparecer aperta para descer a lista
    while not _find_img(f'caixa_{pasta}_email.png', conf=0.9):
        p.press('down')
    
    # seleciona a caixa
    p.press('enter')
    
    # confirmar a seleção
    _click_img('confirma_mover_email.png', conf=0.9)
    time.sleep(10)


@_time_execution
def run():
    
    print('>>> Aguardando documentos...')
    
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Sicalc\\Gerador de guias de DARF WEB\\execução\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    while 0 < 1:
        mes = datetime.datetime.now().month
        ano = datetime.datetime.now().year
        nome_planilha = f'Envia link {mes}-{ano}'
        
        try:
            os.remove(os.path.join('ignore', 'controle', arquivo))
        except:
            pass
        
        status, driver = _initialize_chrome(options)
        driver = login_email(driver)
        time.sleep(3)
        
        print('>>> Aguardando e-mail')
        while re.compile(r'<div class=\"no-items\" style=\"\">Sem itens para mostrar</div>').search(driver.page_source):
            time.sleep(1)
        
        titulo = 'x'
        nao_envia = 'x'
        try:
            driver, titulo, cnpj, cnpj_limpo, vencimento, link_mensagem = captura_link_email(driver)
            
            # determina o tempo de espera entre uma mensagem e outra para tentar evitar span
            numero = random.randint(1, 10)
            time.sleep(numero)
            
            if cnpj_limpo != 'x':
                numero, nome = verifica_o_numero(cnpj_limpo)
            else:
                numero = False
                nome = 'x'
            
            verifica_titulo = titulo.split('-')
            
            emails_diferentes = ['Desconsideração de e', 'Solicitação de documentos ', 'Documentos pendentes de visualização ']
            for diferente in emails_diferentes:
                if verifica_titulo[0] == diferente:
                    nao_envia = 'ok'
                    break
                else:
                    dias = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15',
                            '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
                    for dia in dias:
                        if verifica_titulo[0] == f'Tarefa Boleto sistema de Gestão vencimento dia {dia} concluída ':
                            nao_envia = 'ok'
                            break
                        else:
                            nao_envia = False
                    break
                    
            titulo_sem_emoji = _remove_emojis(titulo)
            
            # se não tiver como enviar
            if nao_envia:
                time.sleep(1)
                mover_email('nao_enviados')
                _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo}', nome=nome_planilha)
                print(f'❗ {titulo}\n')
            
            # se não encontrar o número da planilha
            elif not numero:
                time.sleep(1)
                mover_email('nao_enviados')
                _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo_sem_emoji};Número não encontrado', nome=nome_planilha)
                print('❌ Número não encontrado\n')
            
            # se tiver como enviar e encontrar o número na planilha
            else:
                resultado = envia(numero, titulo, vencimento, link_mensagem)
                # se der erro ao enviar
                if resultado == 'erro':
                    time.sleep(1)
                    mover_email('nao_enviados')
                    _escreve_relatorio_csv(f'{cnpj_limpo};{nome};{numero};{titulo_sem_emoji};Erro ao enviar a mensagem', nome=nome_planilha)
                    print('❌ Erro ao enviar a mensagem\n')
                # se conseguir enviar
                elif resultado == 'ok':
                    time.sleep(1)
                    mover_email('enviados')
                    _escreve_relatorio_csv(f'{cnpj_limpo};{nome};{numero};{titulo_sem_emoji};Mensagem enviada', nome=nome_planilha)
                    print('✔ Mensagem enviada\n')
                # se der algum erro específico ao enviar
                else:
                    time.sleep(1)
                    mover_email('nao_enviados')
                    _escreve_relatorio_csv(f'{cnpj_limpo};{nome};{numero};{titulo_sem_emoji};{resultado}', nome=nome_planilha)
                    print(f'❌ {resultado}\n')
        
        # se der erro em qualquer etapa
        except:
            time.sleep(1)
            print(driver.page_source)
            mover_email('nao_enviados')
            _escreve_relatorio_csv(f'x;x;x;{titulo};Erro ao enviar a mensagem', nome=nome_planilha)
            print('❌ Erro ao enviar a mensagem\n')

        driver.close()
        
            
if __name__ == '__main__':
    run()

