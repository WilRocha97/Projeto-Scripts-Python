# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime, os, time, shutil, re, fitz

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome


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
    
    
def renomear(empresa):
    cnpj, nome = empresa
    download_folder = "V:\\Setor Robô\\Scripts Python\\Veri\\Gerar DARF\\ignore\\docs"
    final_foder = "V:\\Setor Robô\\Scripts Python\\Veri\\Gerar DARF\\execucao\\Guias"
    
    count = 0
    while not os.listdir(download_folder):
        time.sleep(1)
        count += 1
        if count >= 30:
            _escreve_relatorio_csv(f'{cnpj};{nome};Erro, nenhum arquivo baixado')
            print('❌ Erro, nenhum arquivo baixado')
            return False
    
    time.sleep(2)
    for file in os.listdir(download_folder):
        with fitz.open(os.path.join(download_folder, file)) as pdf:
            for page in pdf:
                textinho = page.get_text('text', flags=1 + 2 + 8)
                vencimento = re.compile(r'Pagar este documento até\n(.+)').search(textinho).group(1)

        file_rename = f'{cnpj} - DARF - venc. {vencimento.replace("/", "-")}.pdf'
        shutil.move(os.path.join(download_folder, file), os.path.join(final_foder, file_rename))
        time.sleep(2)
        _escreve_relatorio_csv(f'{cnpj};{nome};Guia gerada no Veri')
        print('✔ Guia gerada no Veri')


def login_veri(empresa, driver):
    driver.get('https://26973312000175.veri-sp.com.br/login')
    print('>>> Acessando o site')
    time.sleep(1)

    dados = "V:\\Setor Robô\\Scripts Python\\SIEG\\Dados.txt"
    f = open(dados, 'r', encoding='utf-8')
    user = f.read()
    user = user.split('/')
    
    # inserir o CNPJ no campo
    driver.find_element(by=By.NAME, value='login').send_keys(user[0])
    time.sleep(1)

    # inserir a senha no campo
    driver.find_element(by=By.NAME, value='senha').send_keys(user[1])
    time.sleep(1)

    # clica em acessar
    driver.find_element(by=By.NAME, value='btn_login').click()
    time.sleep(1)
    
    # gerar a guia de DCTF
    resultado = gerar(empresa, driver)
    return resultado
    

def gerar(empresa, driver):
    cnpj, nome = empresa
    while not localiza_id(driver, 'carouselExampleIndicators'):
        time.sleep(1)
    
    driver.get('https://26973312000175.veri-sp.com.br/dctf_web_nova/index?filter=ATIVA')
    time.sleep(1)
    
    # espera aparecer a barra de pesquisa
    while not localiza_path(driver, '/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input'):
        time.sleep(1)
    
    # insere o cnpj na barra de pesquisa
    driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input').send_keys(cnpj)
    time.sleep(2)
    
    # busca a mensagem de 'Nenhum registro encontrado'
    try:
        nao_encontrada = re.compile(r'class=\"dataTables_empty\">(Nenhum registro encontrado)</td></tr></tbody>').search(driver.page_source).group(1)
    except:
        nao_encontrada = ''
    if nao_encontrada == 'Nenhum registro encontrado':
        _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum registro encontrado')
        print('❗ Nenhum registro encontrado')
        return 'Continue'
    
    # enquanto tiver o botão de gerar
    count = 0
    count2_final = 0
    while localiza_path(driver, '/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[2]/div/div/div[2]/table/tbody/tr[2]/td[11]/div/div/a'):
        # salva a guia se tiver o botão de salvar
        if localiza_path(driver, '/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[2]/div/div/div[2]/table/tbody/tr[2]/td[11]/div/div[3]/a'):
            salvar_guia(driver, empresa)
            return 'Continue'
        
        driver.execute_script("document.querySelector('#example > tbody > tr.odd > td:nth-child(11) > div > div > a').click()")
        time.sleep(2)
        count += 2

        print('>>> Gerando guia...')
        count2 = 0
        while localiza_path(driver, '/html/body/div[6]/div/h2'):
            time.sleep(1)
            count2 += 1
            if count2 >= 60:
                return 'Erro'
        
        count2_final += count2
        while not localiza_path(driver, '/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input'):
            time.sleep(1)
            count += 1

        driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[3]/section/div/div/div/div[2]/div[3]/div/div[1]/div[2]/div/label/input').send_keys(cnpj)
        time.sleep(2)
        count += 2
        if count >= 60 - count2_final:
            return 'Erro'
        
        
def salvar_guia(driver, empresa):
    print('>>> Baixando guia...')
    driver.execute_script("document.querySelector('#example > tbody > tr.odd > td:nth-child(11) > div > div.center > a').click()")
    time.sleep(1)
    renomear(empresa)
    
    
@_time_execution
def run():
    os.makedirs('ignore/docs', exist_ok=True)
    os.makedirs('execucao/Guias', exist_ok=True)
    # p.mouseInfo()

    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Veri\\Gerar DARF\\ignore\\docs",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        cnpj, nome = empresa

        resultado = 'Erro'
        count = 0
        # fazer login no Veri
        while resultado == 'Erro':
            # iniciar o driver do chome
            status, driver = _initialize_chrome(options)
            
            resultado = login_veri(empresa, driver)
            driver.close()
            count += 1
            if count >= 3:
                _escreve_relatorio_csv(f'{cnpj};{nome};Não conseguiu gerar a guia')
                print('❌ Não conseguiu gerar a guia')
                break


if __name__ == '__main__':
    run()
