# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from pyautogui import prompt, confirm
import datetime, os, time, re, shutil, fitz

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers, _barra_de_status
from chrome_comum import _initialize_chrome, _find_by_path, _find_by_id
from pyautogui_comum import _click_img, _find_img


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
    driver.get('https://hub.sieg.com/IriS/#/ComprovantesDePagamentos')

    return driver


def procura_empresa(execucao, tipo, competencia, empresa, driver, options):
    cnpj, nome = empresa
    
    # espera a barra de pesquisa, se não aparecer em 1 minuto, recarrega a página
    timer = 0
    while True:
        if not _find_by_id('select2-ddlCompanyIris-container', driver):
            time.sleep(1)
            timer += 1
            _click_img('comprovantes.png', conf=0.9)
            if timer >= 60:
                print('>>> Tentando novamente\n')
                driver.close()
                status, driver = _initialize_chrome(options)
                driver = login_sieg(driver)
                driver = sieg_iris(driver)
                timer = 0
        else:
            break

    time.sleep(5)
    print('>>> Pesquisando empresa')
    # clica para abrir a barra de pesquisa
    
    if _find_by_id('modalYouTube', driver):
        # Encontre um elemento que esteja fora do modal e clique nele
        elemento_fora_modal = driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div/div/div[2]/div[2]/div[1]/input')
        ActionChains(driver).move_to_element(elemento_fora_modal).click().perform()

    timer = 0
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    while not _find_by_path('/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
        time.sleep(1)
        _click_img('comprovantes.png', conf=0.9)
        if timer >= 60:
            print('>>> Tentando novamente\n')
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)
            driver = sieg_iris(driver)
            timer = 0
        
    # clica para abrir a barra de pesquisa
    driver.find_element(by=By.ID, value='select2-ddlCompanyIris-container').click()
    time.sleep(1)
    
    driver.find_element(by=By.CLASS_NAME, value='select2-search__field').send_keys(cnpj)
    time.sleep(3)
    
    # busca a mensagem de 'Nenhuma empresa encontrada.
    localiza_empresa = re.compile(r'Nenhuma empresa encontrada.').search(driver.page_source)
    if localiza_empresa:
        _escreve_relatorio_csv(f'{cnpj};{nome};Empresa não encontrada', nome=execucao)
        print('❗ Empresa não encontrada')
        return driver
    
    # clica para selecionar a empresa
    driver.find_element(by=By.XPATH, value='/html/body/span/span/span[2]/ul/li').click()
    time.sleep(1)
    
    print('>>> Consultando comprovantes de pagamento')
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    timer = 0
    while not _find_by_path('/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
        time.sleep(1)
        timer += 1
        _click_img('comprovantes.png', conf=0.9)
        
        if re.compile(r'Nenhum dado encontrado').search(driver.page_source) or timer >= 60:
            _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento encontrado para essa empresa', nome=execucao)
            print('❗ Nenhum comprovante de pagamento encontrado para essa empresa')
            return driver
        
    try:
        # clica para expandir a lista de arquivos
        print('>>> Expandindo lista de comprovantes de pagamento')
        driver.find_element(by=By.XPATH, value='/html/body/form/div[5]/div[3]/div[1]/div/div[3]/div/table/tbody/tr[1]/td/div/span/span').click()
        time.sleep(2)
    except:
            re.compile(r'class=\"\">(Nenhum item encontrado!)<').search(driver.page_source).group(1)
            _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento encontrado para essa empresa', nome=execucao)
            print('❗ Nenhum comprovante de pagamento encontrado para essa empresa')
            return driver
    
    pagina = 1
    achou = 'não'
    contador = 0
    erro = ''
    sem_recibo = ''
    # while para tenta percorrer todas as possíveis páginas.
    while True:
        # pega a lista de guias da competência desejada
        if tipo == 'Consulta mensal':
            comprovantes = re.compile(r'/\d\d\d\d</td><td class=\" hidden-sm hidden-xs td-background-dt\">\d\d/' + competencia + '</td><td class=\" hidden-sm hidden-xs td-background-dt\">.+</td><td class=\" col-md-1 hidden-sm hidden-xs td-background-dt\">R\$.+id=\"(.+)\" class=')\
                .findall(driver.page_source)
        else:
            comprovantes = re.compile(r'/\d\d\d\d</td><td class=\" hidden-sm hidden-xs td-background-dt\">\d\d/\d\d/' + competencia + '</td><td class=\" hidden-sm hidden-xs td-background-dt\">.+</td><td class=\" col-md-1 hidden-sm hidden-xs td-background-dt\">R\$.+id=\"(.+)\" class=') \
                .findall(driver.page_source)
        
        # verificação para caso o while anterior não encontre nenhum comprovante
        if achou == 'não':
            # verifica se existe algum comprovante referente a competência digitada
            if not comprovantes:
                _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento referente a {competencia} encontrado para essa empresa', nome=execucao)
                print(f'❗ Nenhum comprovante de pagamento referente a {competencia} encontrado para essa empresa')
                return driver
                
            # verifica se existe algum comprovante referente a competência digitada
            nenhum_dado = re.compile(r'Nenhum dado encontrado').search(driver.page_source)
            if nenhum_dado:
                _escreve_relatorio_csv(f'{cnpj};{nome};Nenhum comprovante de pagamento encontrado para essa empresa', nome=execucao)
                print(f'❗ Nenhum comprovante de pagamento encontrado para essa empresa')
                return driver
        
        # faz o download dos comprovantes
        for comprovante in comprovantes:
            achou = 'sim'
            print('>>> Tentando baixar o comprovantes de pagamento')
            download = True
            while download:
                driver, contador, erro, download = download_comprovante(tipo, contador, driver, competencia, cnpj, comprovante)
            if erro == 'erro':
                sem_recibo = 'Existe um comprovante com o botão de download desabilitado ou com erro ao baixar'
                print(f'❌ Existe um comprovante com o botão de download desabilitado ou com erro ao baixar')
            time.sleep(1)
        
        # tenta ir para a próxima página, se não conseguir sai do while e anota os resultados na planilha
        try:
            driver.find_element(by=By.ID, value='tablePayments_next').click()
            pagina += 1
        except:
            break

        # se clicar no botão para a próxima página aguarda ela carregar.
        while not re.compile(r'paginate_button btn btn-default current\" aria-controls=\"tablePayments\" data-dt-idx=\"' + str(pagina)).search(driver.page_source):
            time.sleep(1)
        
    _escreve_relatorio_csv(f'{cnpj};{nome};Comprovantes {competencia} baixados;{contador} Arquivos;{sem_recibo}', nome=execucao)
    
    time.sleep(1)
    return driver


def download_comprovante(tipo, contador, driver, competencia, cnpj, comprovante):
    if not click(driver,comprovante):
        return driver, contador, 'erro', False
    
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\ignore\\Comprovantes"
    final_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\execução\\Comprovantes"
    
    contador_3 = 0
    contador_4 = 0
    while os.listdir(download_folder) == []:
        if _find_img('varios_arquivos.png', conf=0.9):
            _click_img('confirmar_varios_arquivos.png', conf=0.9)
        if contador_3 > 10:
            if not click(driver, comprovante):
                return driver, contador, 'erro', False
            contador_4 += 1
            contador_3 = 0
        time.sleep(1)
        if contador_4 > 5:
            return driver, contador, 'erro', False
        contador_3 += 1
    
    # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
    for arquivo in os.listdir(download_folder):

        if re.compile(r'\.tmp').search(arquivo):
            timer = 0
            while True:
                print('>>> Aguardando download...')
                time.sleep(1)
                for arq in os.listdir(download_folder):
                    arquivo = arq

                if not re.compile(r'\.tmp').search(arquivo):
                    break
                timer += 1
                if timer > 5:
                    os.remove(os.path.join(download_folder, arquivo))
                    return driver, contador, 'ok', True


        elif re.compile(r'crdownload').search(arquivo):
            timer = 0
            while True:
                print('>>> Aguardando download...')
                time.sleep(1)
                for arq in os.listdir(download_folder):
                    arquivo = arq

                if not re.compile(r'crdownload').search(arquivo):
                    break
                timer += 1
                if timer > 5:
                    os.remove(os.path.join(download_folder, arquivo))
                    return driver, contador, 'ok', True

            
        else:
            if tipo == 'Consulta anual':
                try:
                    with fitz.open(os.path.join(download_folder, arquivo)) as pdf:
                        for page in pdf:
                            textinho = page.get_text('text', flags=1 + 2 + 8)
                            try:
                                competencia = re.compile(r'Data de Vencimento\n.+\n\d\d/(\d\d/\d\d\d\d)').search(textinho).group(1)
                            except:
                                competencia = re.compile(r'Data de Vencimento\n.+\n(\d\d/\d\d\d\d)').search(textinho).group(1)
                except:
                    return driver, contador, 'ok', True
                
            novo_arquivo = arquivo.replace('Pagamento', 'Pagamento_' + competencia.replace('/', '_')).replace('.pdf', f' - {cnpj}.pdf')
            
            while not os.path.isfile(os.path.join(final_folder, novo_arquivo)):
                print('>>> Movendo arquivo')
                try:
                    shutil.move(os.path.join(download_folder, arquivo), os.path.join(final_folder, novo_arquivo))
                    time.sleep(2)
                except:
                    pass

            print(f'✔ {novo_arquivo}')
            contador += 1
            
            for arquivo in os.listdir(download_folder):
                while True:
                    try:
                        os.remove(os.path.join(download_folder, arquivo))
                        break
                    except:
                        pass

    return driver, contador, 'ok', False


def click(driver, comprovante):
    contador_2 = 0
    while True:
        try:
            driver.find_element(by=By.ID, value=comprovante).click()
            break
        except:
            contador_2 += 1
            
        if contador_2 > 5:
            return False
    
    return True


# @_time_execution
@_barra_de_status
def run(window):
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver.set_page_load_timeout(30)
    driver = login_sieg(driver)
    driver = sieg_iris(driver)
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        while True:
            #try:
            time.sleep(2)
            driver.refresh()
            driver = procura_empresa(execucao, tipo, competencia, empresa, driver, options)
            break
        """except:
            driver.close()
            status, driver = _initialize_chrome(options)
            driver = login_sieg(driver)"""

    driver.close()
    

if __name__ == '__main__':
    execucao = 'Download Comprovantes de pagamento IRIS SIEG'
    tipo = confirm(title='Script incrível', buttons=('Consulta mensal', 'Consulta anual'))
    
    if tipo == 'Consulta mensal':
        competencia = prompt(title='Script incrível', text='Qual competência referente?', default='00/0000')
    else:
        competencia = prompt(title='Script incrível', text='Qual competência referente?', default='0000')
    
    os.makedirs('execução/Comprovantes', exist_ok=True)
    os.makedirs('ignore/Comprovantes', exist_ok=True)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download comprovantes de pagamento\\ignore\\Comprovantes",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if empresas:
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is not None:
            run()
