# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import prompt
import os, time, re, csv, shutil

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
    print('>>> Acessando IriS DCTF WEB')
    driver.get('https://hub.sieg.com/IriS/#/ComprovantesDePagamentos')

    return driver


def procura_empresa(competencia, empresa, driver, options):
    cnpj, nome = empresa
    # espera a barra de pesquisa, se não aparecer em 1 minuto, recarrega a página
    timer = 0
    while not localiza_id(driver, 'select2-ddlCompanyIris-container'):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            print('>>> Teantando novamente\n')
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)
            driver = sieg_iris(driver)
            timer = 0

    time.sleep(5)
    print('>>> Pesquisando empresa')
    # clica para abrir a barra de pesquisa
    
    modal = 'sim'
    while modal == 'sim':
        try:
            driver.find_element(by=By.ID, value='select2-ddlCompanyIris-container').click()
            time.sleep(1)
            modal = 'não'
        except:
            try:
                driver.find_element(by=By.CLASS_NAME, value='btn-close-alert').click()
                modal = 'sim'
            except:
                _escreve_relatorio_csv(f'{cnpj};{nome};Erro ao pesquisar empresa')
                print('❗ Erro ao pesquisar empresa')
                return driver
                
    # clica para abrir a barra de pesquisa
    driver.find_element(by=By.CLASS_NAME, value='select2-search__field').send_keys(cnpj)
    time.sleep(3)

    # busca a mensagem de 'Nenhuma empresa encontrada.'
    try:
        re.compile(r'__message\">(Nenhuma empresa encontrada\.)').search(driver.page_source).group(1)
        _escreve_relatorio_csv(f'{cnpj};{nome};Empresa não encontrada')
        print('❗ Empresa não encontrada')
        return driver
    except:
        pass

    # clica para selecionar a empresa
    driver.find_element(by=By.XPATH, value='/html/body/span/span/span[2]/ul/li').click()
    time.sleep(1)
    
    print('>>> Consultando comprovantes de pagamento')
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    contador = 1
    timer = 0
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span'):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento encontrado para esssa empresa')
            print('❗ Nenhum comprovante de pagamento encontrado para esssa empresa')
            return driver
        
    try:
        # clica para expandir a lista de arquivos
        driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span/span').click()
        time.sleep(1.5)
    except:
            re.compile(r'class=\"\">(Nenhum item encontrado!)<').search(driver.page_source).group(1)
            _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento encontrado para esssa empresa')
            print('❗ Nenhum comprovante de pagamento encontrado para esssa empresa')
            return driver
    
    # pega a lista de guias da competência desejada
    comprovantes = re.compile(r'>\d\d/' + competencia + '</td><td class=\" hidden-sm hidden-xs td-background-dt\">-</td><td class=\" col-md-1 hidden-sm hidden-xs td-background-dt\">R\$.+id=\"(.+)\" class=').findall(driver.page_source)
    
    # verifica se existe algum comprovante referênte a competência digitada
    if not comprovantes:
        _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento referênte a {competencia} encontrado para esssa empresa')
        print(f'❗ Nenhum comprovante de pagamento referênte a {competencia} encontrado para esssa empresa')
        return driver
    
    contador = 0
    # faz o download dos comprovantes
    for comprovante in comprovantes:
        driver, contador = download_comrpovante(contador, driver, competencia, cnpj, comprovante)
        if contador == 'erro':
            _escreve_relatorio_csv(f'{cnpj};{nome};Existe um comprovante com o botão de download dezabilitado')
            print(f'❌ Existe um comprovante com o botão de download desabilitado')
            return driver
    
    _escreve_relatorio_csv(f'{cnpj};{nome};Comprovantes baixados;{contador} Arquivos')
    
    time.sleep(1)
    return driver


def download_comrpovante(contador, driver, competencia, cnpj, comprovante):
    try:
        driver.find_element(by=By.ID, value=comprovante).click()
    except:
        return driver, 'erro'
    
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\ignore\\Comprovantes"
    final_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\execução\\Comprovantes"
    while os.listdir(download_folder) == []:
        time.sleep(1)
    
    # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
    for arquivo in os.listdir(download_folder):
        if re.compile(r'crdownload').search(arquivo) or re.compile(r'.tmp').search(arquivo):
            os.remove(os.path.join(download_folder, arquivo))
            driver, contador = download_comrpovante(contador, driver, competencia, cnpj, comprovante)

        else:
            novo_arquivo = arquivo.replace('Pagamento', 'Pagamento_' + competencia.replace('/', '_')).replace('.pdf', f' - {cnpj}.pdf')
            shutil.move(os.path.join(download_folder, arquivo), os.path.join(final_folder, novo_arquivo))
            time.sleep(2)
            print(f'✔ {novo_arquivo}')
            contador += 1
            
    return driver, contador


@_time_execution
def run():
    competencia = prompt(text='Qual competência referênte?', title='Script incrível', default='00/0000')
    os.makedirs('execução/Comprovantes', exist_ok=True)
    os.makedirs('ignore/Comprovantes', exist_ok=True)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\ignore\\Comprovantes",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        erro = 'sim'
        while erro == 'sim':
            try:
                driver = sieg_iris(driver)
                driver = procura_empresa(competencia, empresa, driver, options)
                erro = 'não'
            except:
                try:
                    erro = 'sim'
                    driver.close()
                    status, driver = _initialize_chrome(options)
                    driver = login_sieg(driver)
                except:
                    erro = 'sim'
                    status, driver = _initialize_chrome(options)
                    driver = login_sieg(driver)

    os.remove("V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\execução\\resumo.csv")
        

if __name__ == '__main__':
    run()
