# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from pyautogui import prompt, confirm
from bs4 import BeautifulSoup
import os, time, re, csv, shutil, fitz, pdfkit, pyautogui as p

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
    driver.get('https://hub.sieg.com/IriS/#/DividaAtiva')
    return driver


def consulta_lista(driver, continuar_pagina):
    print('>>> Consultando lista de arquivos\n')
    
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span'):
        time.sleep(1)
    time.sleep(2)
    if localiza_id(driver, 'modalYouTube'):
        # Encontre um elemento que esteja fora do modal e clique nele
        elemento_fora_modal = driver.find_element(by=By.ID, value='txtDataSearch')
        ActionChains(driver).move_to_element(elemento_fora_modal).click().perform()

    time.sleep(1)
    paginas = re.compile(r'>(\d+)</a></span><a class=\"paginate_button btn btn-default next').search(driver.page_source).group(1)
    
    for pagina in range(int(paginas) + 1):
        if pagina == 0:
            continue
        
        if pagina < int(continuar_pagina):
            driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
            # espera a lista de arquivos carregar
            while re.compile(r'id=\"tableActiveDebit_processing\" class=\"\" style=\"display: block;').search(driver.page_source):
                time.sleep(1)
            time.sleep(3)
            continue

        # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
        while not localiza_path(driver, '/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span'):
            time.sleep(1)

        time.sleep(1)
        print(f'>>> Página: {pagina}\n')
        
        info_pagina = driver.page_source
        
        detalhes = re.compile(r'btn-details float-right\" (data-id=\".+\") onclick=\"SeeDetails').findall(driver.page_source)
        for detalhe in detalhes:
            driver.find_element(by=By.XPATH, value=f'//a[@{detalhe}]').click()
            time.sleep(1)
        
        time.sleep(1)
        # pega a lista de guias da competência desejada
        lista_divida = re.compile(r'excel\"><a id=\"(.+)\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"').findall(driver.page_source)
    
        contador = 0
        # faz o download dos comprovantes
        for divida in lista_divida:
            print('>>> Tentando baixar o arquivo')
            download = True
            while download:
                empresa = re.compile(r"col-md-3 td-background-dt hidden-xs\">(.+)</td><td class=\" td-background-dt hidden-xs\">.+excel\"><a id=\"" + divida + "\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"")
                descricao = empresa.search(driver.page_source).group(1)
                    
                driver, contador, erro, download= download_divida(contador, driver, divida, descricao)
                if erro == 'erro':
                    _escreve_relatorio_csv(f'Use a empresa anterior da lista para localizar este arquivo no site;Erro;Erro ao baixar arquivo;Pagina do site: {pagina}')
                    print('>>> Erro ao baixar o arquivo\n')
                    break
        
        if pagina == 25:
            return driver
        
        driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
        # espera a lista de arquivos carregar
        while re.compile(r'id=\"tableActiveDebit_processing\" class=\"\" style=\"display: block;').search(driver.page_source):
            time.sleep(1)
        time.sleep(3)
        
        while info_pagina == driver.page_source:
            print(f'>>> Aguardando Página {pagina}\n')
            time.sleep(2)
        
        time.sleep(1)
        

def espera_pagina(pagina, driver):
    try:
        re.compile(r'paginate_button btn btn-default current\" aria-controls=\"tableActiveDebit\" data-dt-idx=\"' + str(pagina) + '\" tabindex').search(driver.page_source)
        return True
    except:
        return False
    

def download_divida(contador, driver, divida, descricao):
    # formata o texto da descrição da dívida e define as pastas de download do arquivo e a pasta final do PDF
    descricao = descricao.replace('/', ' ')
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa"
    final_folder = f"V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\execução\\Divida Ativa {descricao}"
    
    # cria as pastas de download e a pasta final do arquivo
    os.makedirs(download_folder, exist_ok=True)
    os.makedirs(final_folder, exist_ok=True)
    
    # limpa a pasta de download para não ter conflito com novos arquivos
    for arquivo in os.listdir(download_folder):
        os.remove(os.path.join(download_folder, arquivo))
    
    # tenta clicar em download até conseguir
    while not click(driver, divida):
        time.sleep(5)
    
    # se conseguir, mas não aparecer nada na pasta, continua tentando baixar até o limite de 6 tentativas
    time.sleep(2)
    contador_3 = 0
    contador_4 = 0
    while os.listdir(download_folder) == []:
        if contador_3 > 10:
            while not click(driver, divida):
                time.sleep(5)
            contador_4 += 1
            contador_3 = 0
        if contador_4 > 5:
            return driver, contador, 'erro', True
        time.sleep(2)
        contador_3 += 1
    
    # verifica se os arquivos baixados estão com erro ou ainda não terminaram de baixar
    for arquivo in os.listdir(download_folder):
        # caso exista algum arquivo com problema, tenta de novo o mesmo arquivo
        while re.compile(r'.tmp').search(arquivo):
            try:
                os.remove(os.path.join(download_folder, arquivo))
                return driver, contador, 'ok', True
            except:
                pass
            for arq in os.listdir(download_folder):
                arquivo = arq
        
        while re.compile(r'crdownload').search(arquivo):
            print('>>> Aguardando download...')
            time.sleep(3)
            for arq in os.listdir(download_folder):
                arquivo = arq
        
        else:
            time.sleep(2)
            # se não tiver problema com o arquivo baixado, tenta converter
            for arq in os.listdir(download_folder):
                converte_html_pdf(download_folder, final_folder, arq, descricao)
                time.sleep(2)
                contador += 1
            
            # limpa a pasta de download caso fique algum arquivo nela
            for arquivo in os.listdir(download_folder):
                os.remove(os.path.join(download_folder, arquivo))
                
    return driver, contador, 'ok', False


def click(driver, divida):
    # função para clicar em elementos via ID
    print('>>> Baixando arquivo')
    contador_2 = 0
    clicou = 'não'
    while clicou == 'não':
        try:
            driver.find_element(by=By.ID, value=divida).click()
            clicou = 'sim'
        except:
            contador_2 += 1
            clicou = 'não'
        if contador_2 > 5:
            return False
    
    return True


def converte_html_pdf(download_folder, final_folder, arquivo, descricao):
    # Defina o caminho para o arquivo HTML e PDF
    html_path  = os.path.join(download_folder, arquivo)
    
    # define o novo nome do arquivo PDF
    novo_arquivo = pega_info_arquivo(html_path, descricao)
    
    # Defina o caminho para o arquivo HTML e PDF
    pdf_path = os.path.join(final_folder, novo_arquivo)
    
    # Localização do executável do wkhtmltopdf
    wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

    # Configuração do pdfkit para utilizar o wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    # tenta converter o PDF, se não conseguir ano na planilha o nome do arquivo em html que deu erro na conversão
    # e coloca ele em uma pasta separada para arquivos em html
    try:
        time.sleep(2)
        pdfkit.from_file(html_path, pdf_path, configuration=config)
        time.sleep(2)
    except:
        _escreve_relatorio_csv(f'{novo_arquivo.replace(".pdf", ".html")};Arquivo HTML salvo;Erro ao converter arquivo')
        final_folder = 'V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\execução\\Arquivos em HTML'
        os.makedirs(final_folder, exist_ok=True)
        shutil.move(html_path, os.path.join(final_folder, novo_arquivo.replace('.pdf', '.html')))
        print(f'❗ Erro ao converter arquivo\n')
        time.sleep(2)
        return False
    
    # anota o arquivo que acabou de baixar e converter
    print(f'✔ {novo_arquivo}\n')
    _escreve_relatorio_csv(f'{novo_arquivo.replace(" - ", ";").replace("-", ";").replace("EXTINTA POR PRESCRICAO;ROTINA AUTOMATICA", "EXTINTA POR PRESCRICAO-ROTINA AUTOMATICA")}')
    return True
    

def pega_info_arquivo(html_path, descricao):
    # abrir o arquivo html
    with open(html_path, 'r', encoding='utf-8') as file:
        html = file.read()
    
        # extrair o texto do arquivo html usando BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text()
        
        # pega as infos da empresa e a inscrição do documento porque tem empresas com mais de um arquivo
        try:
            try:
                nome = re.compile(r'RFB\n\n\n\nNome:\n(.+)').search(text).group(1)
            except:
                nome = re.compile(r'Devedor Principal:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
            cpf_cnpj = re.compile(r'CNPJ/CPF:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
            inscricao = re.compile(r'Inscrição:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
        except:
            nome = re.compile(r'Devedor Principal:\n(.+)\n').search(text).group(1)
            cpf_cnpj = re.compile(r'CNPJ/CPF:\n(.+)\n').search(text).group(1)
            inscricao = re.compile(r'N\.º Inscrição:\n(.+)\n').search(text).group(1)
        
        # formata os textos
        inscricao = inscricao.replace('.', '').replace('-', '').replace(' ', '')
        nome = nome[:20]
        nome = nome.replace('/', '').replace("-", "")
        cpf_cnpj = cpf_cnpj.replace('.', '').replace('/', '').replace('-', '')
        
        nome_pdf = f'{nome} - {cpf_cnpj} - {descricao}-{inscricao}.pdf'
        nome_pdf = nome_pdf.replace('  ', ' ')
    
    return nome_pdf


@_time_execution
def run():
    continuar_pagina = p.prompt(text='Continuar a partir de qual página?')
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": False  # It will not show PDF directly in chrome
    })
    
    # iniciar o driver do chome
    status, driver = _initialize_chrome(options)
    driver = login_sieg(driver)
    driver = sieg_iris(driver)
    driver = consulta_lista(driver, continuar_pagina)

    driver.close()
    

if __name__ == '__main__':
    run()
