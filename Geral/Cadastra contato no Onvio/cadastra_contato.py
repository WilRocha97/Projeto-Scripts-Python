import time, re, os
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _find_by_id, _find_by_path, _initialize_chrome
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Onvio.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


def login(driver):
    print(f'>>> Logando no Onvio')
    # abre a tela de login
    driver.get('https://onvio.com.br/#/')
    while not _find_by_id('trauth-continue-signin-btn', driver):
        time.sleep(1)
    
    # clica para iniciar o login
    driver.execute_script("document.getElementById('trauth-continue-signin-btn').click()")
    
    # aguarda o campo de email carregar
    while not _find_by_id('username', driver):
        time.sleep(1)
    # insere o email e confirma
    driver.find_element(by=By.ID, value='username').send_keys(user[0])
    driver.find_element(by=By.XPATH, value='/html/body/div/div/div/div[2]/div/main/section/div/div/div/div[1]/div/form/div[2]/button').click()
    
    # aguarda o campo de senha carregar
    while not _find_by_id('password', driver):
        time.sleep(1)
    # insere a senha e confirma
    driver.find_element(by=By.ID, value='password').send_keys(user[1])
    driver.find_element(by=By.XPATH, value='/html/body/div/div/div/div[2]/div/main/section/div/div/div/form/div[3]/button').click()
    
    # aguarda o site carregar
    while not _find_by_id('mainView', driver):
        time.sleep(1)
    
    return driver

    
def adiciona_cliente(driver, nome_contato, numero):
    print(f'>>> Abrindo o Messenger')
    # abre a home do messenger
    driver.get('https://onvio.com.br/br-messenger/home')
    
    # clica para abrir o messenger
    while True:
        try:
            driver.find_element(by=By.XPATH, value='/html/body/app-root/div[1]/on-header/bm-custom-header/bento-off-canvas-menu/div[5]/div/on-nav/nav/div[2]').click()
            break
        except:
            time.sleep(1)
    
    # o messenger abre em outra aba
    # aguarda dois segundos para abrir uma nova aba, troca o driver para essa nova aba e a fecha, depois volta para a aba inicial
    # esse procedimento serve para, gerar cookies da janela de contatos do messenger e o script conseguir abri-la diretamente na mesma aba
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    
    # abre diretamente na aba de cadastro de contato
    driver.get('https://app.gestta.com.br/messenger-v2/#/chat/contact-list')
    
    print(f'>>> Preenchendo dados do cliente')
    while not _find_by_path('/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/button', driver):
        time.sleep(1)
    # clica para abrir o formulário
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/button').click()
    time.sleep(1)
    
    # insere as informações do contato: nome, DDD e número
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/div/div/div/form/div[1]/div[1]/div[1]/div[2]/input').send_keys(nome_contato)
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/div/div/div/form/div[1]/div[2]/div[2]/div[1]/div[2]/input').send_keys(numero[:2])
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/div/div/div/form/div[1]/div[2]/div[3]/div[1]/div[2]/input').send_keys(numero[2:])
    time.sleep(1)
    
    # confirma o novo contato
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div[4]/div/div/div/div/div/form/div[3]/div/div[2]/button').click()
    time.sleep(1)
    
    return driver
    
    
def verificacoes(driver):
    print('>>> Verificando contato')
    print(driver.page_source)
    if re.compile(r'Este número de telefone já está cadastrado nos contatos').search(driver.page_source):
        return False, driver, 'Este número de telefone já está cadastrado nos contatos.'
    
    return True, driver, 'ok'


def manda_primeira_mensagem(driver):
    # primeira mensagem enviada ao cliente.
    primeira_mensagem = ('Prezado cliente.\n'
                         'Com o intuito de proporcionar um atendimento com mais qualidade, agilidade e praticidade, a Veiga & Postal disponibiliza mais um canal de WhatsApp, sendo o número (19) 9 8146-6111.\n'
                         'Além disso, será através deste novo número que a Veiga & Postal enviará informativos a sua empresa e relatório de pendências tributárias, caso tenha.\n'
                         'Em caso de dúvida estamos à disposição.')
    
    while not _find_by_path('/html/body/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/div[6]/div[4]/div/textarea', driver):
        time.sleep(1)
    
    # insere a mensagem no campo
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/div[6]/div[4]/div/textarea').send_keys(primeira_mensagem)
    # clica em enviar mensagem
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div[2]/div[2]/div/div[3]/div[2]/div[6]/div[4]/div/div[2]').click()
    
    time.sleep(5)
    
    if re.compile(r'Este arquivo está sendo enviado para um número inexistente, verifique o número do contato.').search(driver.page_source):
        return False, driver, 'Este arquivo está sendo enviado para um número inexistente, verifique o número do contato.'
    
    print(driver.page_source)
    return True, driver, 'Número cadastrado, mensagem enviada'

    

@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o índice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    status, driver = _initialize_chrome(options)
    # faz o login uma vez
    driver = login(driver)
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, razao, nome, numero = empresa
        
        # printa o índice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        nome_contato = nome + ' - ' + razao
        driver = adiciona_cliente(driver, nome_contato, numero)
        
        resultado, driver, mensagem = verificacoes(driver)
        if not resultado:
            _escreve_relatorio_csv(f'{cnpj};{razao};{nome};{numero};{mensagem}')
            print(f'❌ {mensagem}')
            continue
        
        resultado, driver, mensagem = manda_primeira_mensagem(driver)
        _escreve_relatorio_csv(f'{cnpj};{razao};{nome};{numero};{mensagem}')
        print(f'✔ {mensagem}')


if __name__ == '__main__':
    run()
