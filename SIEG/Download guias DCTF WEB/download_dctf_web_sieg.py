# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from pyautogui import prompt
import os, time, re, csv, shutil

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path


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
    

def sieg_iris(driver):
    print('>>> Acessando IriS DCTF WEB')
    driver.get('https://hub.sieg.com/IriS/#/DCTFWeb')

    return driver


def procura_empresa(competencia, empresa, driver, options):
    cnpj, nome = empresa
    # espera a barra de pesquisa, se não aparecer em 1 minuto, recarrega a página
    timer = 0
    while not _find_by_id('select2-ddlCompanyIris-container', driver):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            print('>>> Tentando novamente\n')
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)
            driver = sieg_iris(driver)
            timer = 0
        
    time.sleep(5)
    print('>>> Pesquisando empresa')

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

    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    timer = 0
    while not _find_by_path('/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            try:
                re.compile(r'class=\"\">(Nenhum item encontrado!)<').search(driver.page_source).group(1)
                _escreve_relatorio_csv(f'{cnpj};{nome};Nenhuma guia de pagamento encontrado para essa empresa')
                print('❗ Nenhuma guia de pagamento encontrado para essa empresa')
                return driver
            except:
                print('>>> Tentando novamente\n')
                driver.close()
                status, driver = _initialize_chrome(options)
                driver = login_sieg(driver)
                driver = sieg_iris(driver)
                procura_empresa(competencia, empresa, driver, options)
                return driver
    
    print('>>> Verificando download')
    try:
        # clica para expandir a lista de arquivos
        driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span/span').click()
        time.sleep(2)
    except:
        re.compile(r'class=\"\">(Nenhum item encontrado!)<').search(driver.page_source).group(1)
        _escreve_relatorio_csv(f'{cnpj};{nome};Nenhuma guia de pagamento encontrado para essa empresa')
        print('❗ Nenhuma guia de pagamento encontrado para essa empresa')
        return driver
    
    # pega a lista de guias da competência desejada
    guias = re.compile(r'</td><td class=\"td-background-dt sorting_2\">' + competencia + '</td><td class=\" td-background-dt\">.+</td><td class=\" td-background-dt\">R\$&nbsp;.+</td><td class=\" td-background-dt hidden-sm hidden-xs\">.+\n.+\n.+<a id=\"(.+)\" class=\"btn iris-') \
        .findall(driver.page_source)

    # verifica se existe alguma guia referente a competência digitada
    if not guias:
        _escreve_relatorio_csv(f'{cnpj};{nome};Nenhuma guia de pagamento referente a {competencia} encontrado para essa empresa')
        print(f'❗ Nenhuma guia de pagamento referente a {competencia} encontrado para essa empresa')
        return driver

    contador = 0
    erro = ''
    sem_guia = ''
    # faz o download dos comprovantes
    for guia in guias:
        download = True
        while download:
            driver, contador, erro, download = download_comprovante(contador, driver, cnpj, guia)
        if erro == 'erro':
            sem_guia = 'Existe uma guia com o botão de download desabilitado'
            print(f'❌ Existe uma guia com o botão de download desabilitado')
            _escreve_relatorio_csv(f'{cnpj};{nome};{sem_guia}')
            return driver

    _escreve_relatorio_csv(f'{cnpj};{nome};Guia {competencia} baixada;{contador} Arquivos', nome='Download DCTF web')
    
    # deleta o andamento que já baixou a guia
    file = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Dados.csv"
    with open(file, 'r', encoding='latin-1') as f:
        dados = f.readlines()

    dados = list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))

    f = open(file, 'w', encoding='latin-1')
    f.write('')
    f.close()

    for linha in dados:
        if linha != empresa:
            cnpj, nome = linha
            f = open(file, 'a', encoding='latin-1')
            f.write(f'{cnpj};{nome}\n')
            f.close()
            
    time.sleep(1)
    return driver
    
def download_comprovante(contador, driver, cnpj, guia):
    if not click(driver, cnpj, guia):
        return driver, contador, 'erro', False
    
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Guias"
    final_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\Guias"

    print('>>> Aguardando o download da guia')
    contador_3 = 0
    while os.listdir(download_folder) == []:
        if contador_3 > 20:
            if not click(driver, cnpj, guia):
                return driver, contador, 'erro', False
            contador_3 = 0
        time.sleep(1)
        contador_3 += 1
    
    # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
    for arquivo in os.listdir(download_folder):
        if re.compile(r'.tmp').search(arquivo):
            os.remove(os.path.join(download_folder, arquivo))
            return driver, contador, 'ok', True
        
        while re.compile(r'crdownload').search(arquivo):
            print('>>> Aguardando download...')
            time.sleep(3)
            for arq in os.listdir(download_folder):
                arquivo = arq
        
        else:
            while not os.path.isfile(os.path.join(final_folder, arquivo)):
                print('>>> Movendo arquivo')
                shutil.move(os.path.join(download_folder, arquivo), os.path.join(final_folder, arquivo))
                time.sleep(2)
                        
            print(f'✔ {arquivo}')
            contador += 1
            
            for arquivo in os.listdir(download_folder):
                os.remove(os.path.join(download_folder, arquivo))
    
    return driver, contador, 'ok', False


def click(driver, cnpj, guia):
    codigo_id = guia.split('-')[1]
    contador_2 = 0
    clicou = 'não'
    while clicou == 'não':
        try:
            driver.find_element(by=By.ID, value=guia).click()
            time.sleep(0.5)
            nome = "DownloadDCTFWeb('" + codigo_id + "', '" + cnpj + "', 'DarfOrDae')"
            driver.find_element(by=By.CSS_SELECTOR, value='a[onclick^="' + nome + '"]').click()
            clicou = 'sim'
            print('>>> Baixando guia')
        except:
            contador_2 += 1
            clicou = 'não'
        if contador_2 > 5:
            return False
    
    return True


@_time_execution
def run():
    competencia = prompt(text='Qual competência referente?', title='Script incrível', default='00/0000')
    os.makedirs('execução/Guias', exist_ok=True)
    os.makedirs('ignore/Guias', exist_ok=True)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Guias",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    contador = 1
    while True:
        print(f'\n\nIniciando rotina Nº {contador} ----------------------------------------------------------\n')
        # abre a planilha de dados
        file = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\ignore\\Dados.csv"
        with open(file, 'r', encoding='latin-1') as f:
            dados = f.readlines()

        empresas = list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))

        # configurar um indice para a planilha de dados
        if contador == 1:
            index = _where_to_start(tuple(i[0] for i in empresas))
            if index is None:
                return False
        
        else:
            index = 0
        
        # iniciar o driver do chome
        status, driver = _initialize_chrome(options)
        driver = login_sieg(driver)
        
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
    
            # configurar o indice para localizar em qual empresa está
            _indice(count, total_empresas, empresa, index)
            while True:
                try:
                    driver = sieg_iris(driver)
                    driver = procura_empresa(competencia, empresa, driver, options)
                    break
                except:
                    try:
                        driver.close()
                        status, driver = _initialize_chrome(options)
                        driver = login_sieg(driver)
                    except:
                        status, driver = _initialize_chrome(options)
                        driver = login_sieg(driver)

        contador += 1
        os.remove("V:\\Setor Robô\\Scripts Python\\SIEG\\Download guias DCTF WEB\\execução\\resumo.csv")
        

if __name__ == '__main__':
    run()
