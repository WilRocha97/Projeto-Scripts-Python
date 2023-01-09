# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, re, csv

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
    

def login_sieg(driver):
    driver.get('https://auth.sieg.com/')
    print('>>> Acessando o site')
    time.sleep(1)
    
    # inserir o CNPJ no campo
    driver.find_element(by=By.ID, value='txtEmail').send_keys('willian.rocha@veigaepostal.com.br')
    time.sleep(1)
    
    # inserir a senha no campo
    driver.find_element(by=By.ID, value='txtPassword').send_keys('Milenniumfalcon9')
    time.sleep(1)
    
    # clica em acessar
    driver.find_element(by=By.ID, value='btnSubmit').click()
    time.sleep(1)
    
    return driver
    

def sieg_iris(driver):
    driver.get('https://hub.sieg.com/IriS/#/DCTFWeb')
    print('>>> Acessando IriS DCTF WEB')

    return driver


def procura_empresa(empresa, driver, options):
    cnpj, nome = empresa
    # espera a barra de pesquisa, se não aparecer em 1 minuto, recarrega a página
    timer = 0
    while not localiza_id(driver, 'select2-ddlCompanyIris-container'):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)
            driver = sieg_iris(driver)
        
    time.sleep(5)
    # clica para abrir a barra de pesquisa
    driver.find_element(by=By.ID, value='select2-ddlCompanyIris-container').click()
    time.sleep(1)

    # clica para abrir a barra de pesquisa
    driver.find_element(by=By.CLASS_NAME, value='select2-search__field').send_keys(cnpj)
    time.sleep(3)

    # busca a mensagem de 'Nenhuma empresa encontrada.'
    try:
        nao_encontrada = re.compile(r'__message\">(Nenhuma empresa encontrada\.)').search(driver.page_source).group(1)
    except:
        nao_encontrada = ''
        
    if nao_encontrada == 'Nenhuma empresa encontrada.':
        _escreve_relatorio_csv(f'{cnpj};{nome};Empresa não encontrada')
        print('❗ Empresa não encontrada')
        return driver

    # clica para selecionar a empresa
    driver.find_element(by=By.XPATH, value='/html/body/span/span/span[2]/ul/li').click()
    time.sleep(1)

    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div[1]/div/div[4]/div/table/tbody/tr[2]/td[9]/div/a[1]'):
        time.sleep(1)
    
    print('>>> Verificando donwload')
    try:
        # clica em download
        driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div[1]/div/div[4]/div/table/tbody/tr[2]/td[9]/div/a[1]').click()
        time.sleep(1)
        # download dctf
        driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div[1]/div/div[4]/div/table/tbody/tr[2]/td[9]/div/ul/li[3]/a').click()
        time.sleep(1)
    
        data = re.compile(r'td-background-dt hidden-md hidden-sm hidden-xs\">(.+)</td><td class=\"td-background-dt sorting_2\">(\d\d/\d\d\d\d)').search(driver.page_source).group(2)
        data = data.replace('/', '_')
        data2 = data.split('_')
        
        download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\Guias"
        guia = os.path.join(download_folder, cnpj + '_DCTFWEB_DarfOrDae_' + data + '.pdf')
        guia2 = os.path.join(download_folder, cnpj + '_DCTFWEB_DarfOrDae_' + str(data2[1]) + '.pdf')
        while not os.path.exists(guia):
            if os.path.exists(guia2):
                break
            time.sleep(1)
        
        data = data.replace('_', '-')
        _escreve_relatorio_csv(f'{cnpj};{nome};Guia salva {data}', nome='Download DCTF web')
        
        # deleta o andamento que já baixou a guia
        file = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Dados.csv"
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()
        
        dados  = list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))
    
        f = open(file, 'w', encoding='latin-1')
        f.write('')
        f.close()
        
        for linha in dados:
            if linha != empresa:
                cnpj, nome = linha
                f = open(file, 'a', encoding='latin-1')
                f.write(f'{cnpj};{nome}\n')
                f.close()
            
        print('✔ Download da guia concluído')
    except:
        _escreve_relatorio_csv(f'{cnpj};{nome};Não possuí guia para download')
        print('❌ Não possuí guia para download')
        return driver
            
    return driver


@_time_execution
def run():
    os.makedirs('execução/Guias', exist_ok=True)

    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    while 1 < 2:
        # abre a planilha de dados
        file = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Dados.csv"
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()

        empresas = list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))

        # configurar um indice para a planilha de dados
        index = 0

        # iniciar o driver do chome
        status, driver = _initialize_chrome(options)
        driver = login_sieg(driver)
        driver = sieg_iris(driver)
        
        erro = 'sim'
        while erro == 'sim':
            try:
                total_empresas = empresas[index:]
                for count, empresa in enumerate(empresas[index:], start=1):
            
                    # configurar o indice para localizar em qual empresa está
                    _indice(count, total_empresas, empresa)
                    
                    driver = procura_empresa(empresa, driver, options)
                    
                erro = 'não'
                driver.close()
            except:
                try:
                    erro = 'sim'
                    driver.close()
                except:
                    pass
            
    
if __name__ == '__main__':
    run()
