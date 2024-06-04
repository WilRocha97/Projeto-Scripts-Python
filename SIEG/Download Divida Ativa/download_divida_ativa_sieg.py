# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import os, time, re, shutil, fitz, pdfkit, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome, _find_by_path, _find_by_id

ultima_divida = 'ultima_divida.txt'


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
    f = open(dados, 'r', encoding='latin-1')
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


def consulta_lista(driver, andamentos):
    ud = open(ultima_divida, 'r', encoding='utf-8').read()
    if ud != '':
        try:
            continuar_pagina = ud.split('|')[0]
            cpf_cnpj_anterior = ud.split('|')[1]
            processo_anterior = ud.split('|')[2]
            inscricao_anterior = ud.split('|')[3]
            print(f'Última divida salva:\n'
                  f'Página: {continuar_pagina}\n'
                  f'CPF/CNPJ: {cpf_cnpj_anterior}\n'
                  f'Número do processo: {processo_anterior}\n'
                  f'Número da inscrição: {inscricao_anterior}\n\n')

        except:
            continuar_pagina = ud
            cpf_cnpj_anterior = 0
            processo_anterior = 0
            inscricao_anterior = 0
    else:
        continuar_pagina = 1
        cpf_cnpj_anterior = 0
        processo_anterior = 0
        inscricao_anterior = 0
        
    print('>>> Consultando lista de arquivos\n')
    
    # espera a lista de arquivos carregar, se não carregar tenta pesquisar novamente
    timer = 0
    while not _find_by_path('/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
        time.sleep(1)
        timer += 1
        if timer >= 60:
            driver.close()
            return False
        
    time.sleep(2)
    if _find_by_id('modalYouTube', driver):
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
        while not _find_by_path('/html/body/form/div[5]/div[3]/div/div/div[3]/div/table/tbody/tr[1]/td/div/span', driver):
            time.sleep(1)

        time.sleep(1)
        print(f'>>> Página: {pagina}\n')
        with open(ultima_divida, 'w', encoding='utf-8') as f:
            f.write(str(pagina))
            
        info_pagina = driver.page_source
        
        detalhes = re.compile(r'btn-details float-right\" (data-id=\".+\") onclick=\"SeeDetails').findall(driver.page_source)
        print('Abrindo listas das empresas...')
        for detalhe in detalhes:
            pixels = 1
            while True:
                try:
                    print(f'ID da lista: {detalhe}')
                    driver.find_element(by=By.XPATH, value=f'//a[@{detalhe}]').click()
                    time.sleep(0.1)
                    break
                except:
                    driver.execute_script(f"window.scrollTo(0, {pixels})")
                    pixels += 1
                    pass
                
        time.sleep(1)
        # pega a lista de guias da competência desejada
        lista_divida = re.compile(r'excel\"><a id=\"(.+)\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"').findall(driver.page_source)
        
        nome = False
        contador = 0
        # faz o download dos comprovantes
        for count, divida in enumerate(lista_divida):
            for i in range(100):
                id_empresa = re.compile(r'<span class=\"title-datatable\">(.+)\((\d+)\)<span(.+\n){' + str(i) + '}.+col-md-2 td-background-dt\">(.+)</td><td class=\" td-background-dt\">(.+)</td><td class=\" td-background-dt hidden-sm hidden-xs\">(\d.+)</td><td class=\" td-background-dt hidden-sm hidden-xs\">.+excel\"><a id=\"' + divida).search(driver.page_source)
                if not id_empresa:
                    continue
                nome = id_empresa.group(1).replace('&amp;', '&')
                cpf_cnpj = id_empresa.group(2)
                tipo = id_empresa.group(4)
                processo = id_empresa.group(5)
                inscricao = id_empresa.group(6)
                break
            
            if not nome:
                print(driver.page_source)
                print(divida)
                raise Exception('Rapadura é doce mas não é mole não')
            
            # verifica qual foi a última empresa consultada da lista
            if cpf_cnpj_anterior != 0:
                # se achar a última dívida salva da execução anterior, pule mais um se não irá baixar a mesma dívida novamente
                if not cpf_cnpj == cpf_cnpj_anterior and not processo == processo_anterior and not inscricao == inscricao_anterior:
                    continue
                cpf_cnpj_anterior = 0
                continue
            
            print('>>> Tentando baixar o arquivo')
            download = True
            while download:
                empresa = re.compile(r'col-md-3 td-background-dt hidden-xs\">(.+)</td><td class=\" td-background-dt hidden-xs\">.+excel\"><a id=\"' + divida + r'\" class=\"btn iris-btn iris-btn-orange iris-btn-sm margin-left\"')
                try:
                    descricao = empresa.search(driver.page_source).group(1)
                except:
                    descricao = 'SEM DESCRIÇÃO'

                driver, contador, erro, download = download_divida(contador, driver, divida, descricao, tipo, processo, inscricao, andamentos)
                if erro == 'erro':
                    try:
                        _escreve_relatorio_csv(f'{cpf_cnpj};{nome};{tipo};{processo};{inscricao};Erro ao baixar o arquivo', nome=andamentos)
                        print('>>> Erro ao baixar o arquivo\n')
                    except:
                        print(driver.page_source)
                        raise Exception(f'Então tá bom {divida}')
                    break
            
            with open(ultima_divida, 'w', encoding='utf-8') as f:
                f.write(f'{pagina}|{cpf_cnpj}|{processo}|{inscricao}')
                
        if pagina == 25:
            return driver
        
        # próxima página
        try:
            driver.find_element(by=By.ID, value='tableActiveDebit_next').click()
        except:
            return 'final'
            
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
    

def download_divida(contador, driver, divida, descricao, tipo, processo, inscricao, andamentos):
    # formata o texto da descrição da dívida e define as pastas de download do arquivo e a pasta final do PDF
    descricao = descricao.replace('/', ' ')
    download_folder = "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa"
    final_folder = f"V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\execução\\Divida Ativa"
    
    # cria as pastas de download e a pasta final do arquivo
    os.makedirs(download_folder, exist_ok=True)
    
    # limpa a pasta de download para não ter conflito com novos arquivos
    for arquivo in os.listdir(download_folder):
        os.remove(os.path.join(download_folder, arquivo))
    
    # tenta clicar em download até conseguir
    while not click(driver, divida):
        time.sleep(5)
    
    # se conseguir, mas não aparecer nada na pasta, continua tentando baixar até o limite de 6 tentativas
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
            # se não tiver problema com o arquivo baixado, tenta converter
            for arq in os.listdir(download_folder):
                resultado = converte_html_pdf(download_folder, final_folder, arq, descricao, tipo, processo, inscricao, andamentos)
                if resultado == 'erro':
                    return driver, contador, 'erro', False
                time.sleep(2)
                contador += 1
            
                # limpa a pasta de download caso fique algum arquivo nela
                os.remove(os.path.join(download_folder, arq))
                
    return driver, contador, 'ok', False


def click(driver, divida):
    # função para clicar em elementos via ID
    print('>>> Baixando arquivo')
    contador_2 = 0
    while True:
        try:
            driver.find_element(by=By.ID, value=divida).click()
            break
        except:
            contador_2 += 1
            
        if contador_2 > 5:
            return False
    
    return True


def converte_html_pdf(download_folder, final_folder, arquivo, descricao, tipo, processo, inscricao, andamentos):
    # Defina o caminho para o arquivo HTML e PDF
    html_path = os.path.join(download_folder, arquivo)
    
    # Estrutura básica do HTML que queremos adicionar
    estrutura_inicial = """<!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
    </head>
    <body>
    """
    
    while True:
        try:
            # Leitura do conteúdo do arquivo existente
            with open(html_path, 'r', encoding='utf-8') as file:
                conteudo_existente = file.read()
            break
        except PermissionError:
            pass
            
    # Adicionando a estrutura inicial ao conteúdo existente
    novo_conteudo = estrutura_inicial + conteudo_existente + "\n</body>\n</html>"
    while True:
        try:
            # Salvando o conteúdo modificado de volta ao arquivo
            with open(html_path, 'w', encoding='utf-8') as file:
                file.write(novo_conteudo)
            break
        except PermissionError:
            pass
    
    # define o novo nome do arquivo PDF
    novo_arquivo, descricao, nome, cpf_cnpj = pega_info_arquivo(html_path, descricao)
    if not novo_arquivo:
        return 'erro'
    
    # coloca na pasta de ativas ou extintas
    if re.compile('EXTINTA').search(descricao) or re.compile('NEGOCIAD').search(descricao) or re.compile('Parcelad').search(descricao):
        final_folder = f'{final_folder} EXTINTAS OU NEGOCIADAS'
    else:
        final_folder = f'{final_folder} ATIVAS'
    os.makedirs(final_folder, exist_ok=True)
    
    # Defina o caminho para o arquivo HTML e PDF
    pdf_path = os.path.join(final_folder, novo_arquivo)
    
    # Localização do executável do wkhtmltopdf
    wkhtmltopdf_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

    # Configuração do pdfkit para utilizar o wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    
    options = {
        'page-size': 'Letter',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
        'encoding': "UTF-8",
        'custom-header': [
            ('Accept-Encoding', 'gzip')
        ],
        'cookie': [
            ('cookie-empty-value', '""'),
            ('cookie-name1', 'cookie-value1'),
            ('cookie-name2', 'cookie-value2')
        ],
        'no-outline': None
    }
    
    # tenta converter o PDF, se não conseguir anota na planilha o nome do arquivo em html que deu erro na conversão
    # e coloca ele em uma pasta separada para arquivos em html
    try:
        pdfkit.from_file(html_path, pdf_path, configuration=config, options=options)
    except:
        _escreve_relatorio_csv(f'{cpf_cnpj};{nome};{tipo};{processo};{inscricao};Erro ao converter arquivo, arquivo HTML salvo: {novo_arquivo.replace(".pdf", ".html")}', nome=andamentos)
        final_folder = 'V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\execução\\Arquivos em HTML'
        os.makedirs(final_folder, exist_ok=True)
        shutil.move(html_path, os.path.join(final_folder, novo_arquivo.replace('.pdf', '.html')))
        print(f'❗ Erro ao converter arquivo\n')
        return False
    
    # anota o arquivo que acabou de baixar e converter
    print(f'✔ {novo_arquivo}\n')
    _escreve_relatorio_csv(f'{cpf_cnpj};{nome};{tipo};{processo};{inscricao};{descricao}', nome=andamentos)
    return True
    

def pega_info_arquivo(html_path, descricao):
    # abrir o arquivo html
    while True:
        print('>>> Tentando abrir o novo arquivo html...')
        try:
            # Salvando o conteúdo modificado de volta ao arquivo
            with open(html_path, 'r', encoding='utf-8') as file:
                html = file.read()
            break
        except PermissionError:
            pass
    
        # extrair o texto do arquivo html usando BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text()
        
        # pega as infos da empresa e a inscrição do documento porque tem empresas com mais de um arquivo
        if re.compile(r'Tempo de conexão esgotado').search(text):
            return False, 'Erro ao acessar o arquivo no portal REGULARIZE', False, False
            
        try:
            nome_inteiro = re.compile(r'RFB\n\n\n\nNome:\n(.+)').search(text).group(1)
        except:
            try:
                nome_inteiro = re.compile(r'Devedor Principal:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
            except:
                try:
                    nome_inteiro = re.compile(r'Devedor principal: (.+)/CNPJ:').search(text).group(1)
                except:
                    try:
                        nome_inteiro = re.compile(r'Devedor Principal:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        try:
            cpf_cnpj = re.compile(r'CNPJ/CPF:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
        except:
            try:
                cpf_cnpj = re.compile(r'CNPJ:.+(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)').search(text).group(1)
            except:
                try:
                    cpf_cnpj = re.compile(r'CNPJ:.+(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(text).group(1)
                except:
                    try:
                        cpf_cnpj = re.compile(r'CNPJ/CPF:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        try:
            inscricao = re.compile(r'Inscrição:\n.+\n\n.+ {2}(.+)\n').search(text).group(1)
        except:
            try:
                inscricao = re.compile(r'Nº inscrição:\s+(\d.+\d).+tuação da inscrição').search(text).group(1)
            except:
                try:
                    sem_inscricao = re.compile(r'Nº inscrição:\s+Situação da inscrição').search(text)
                    if sem_inscricao:
                        inscricao = 'Não foi possível recuperar as informações da inscrição'
                    else:
                        raise Exception
                except:
                    try:
                        inscricao = re.compile(r'N\.º Inscrição:\n(.+)\n').search(text).group(1)
                    except:
                        print(text)
                        raise Exception('Errei fui muleke')
        
        if descricao == 'SEM DESCRIÇÃO':
            try:
                descricao = re.compile(r'Situação da inscrição:\s\s(.+)\s\sInformações gerais').search(text).group(1)
            except:
                try:
                    descricao = re.compile(r'Situação da inscrição: Benefício Fiscal - (.+)\s\sInformações gerais').search(text).group(1)
                except:
                    print(text)
                    raise Exception('Errei fui muleke')

        # formata os textos
        nome_inteiro = nome_inteiro.replace('&amp;', '&')
        inscricao = inscricao.replace('.', '').replace('-', '').replace(' ', '')
        nome = nome_inteiro[:20]
        nome = nome.replace('/', '').replace("-", "")
        cpf_cnpj = cpf_cnpj.replace('.', '').replace('/', '').replace('-', '')
        
        nome_pdf = f'{nome} - {cpf_cnpj} - {descricao}-{inscricao}.pdf'
        nome_pdf = nome_pdf.replace('  ', ' ').replace('/', '-')
    
    return nome_pdf, descricao, nome_inteiro, cpf_cnpj


@_time_execution
def run():
    andamentos = 'Consulta Divida Ativa SIEG IRIS'
    continuar = p.confirm(text='Continuar consulta anterior?', buttons=('Sim', 'Não'))
    if continuar == 'Não':
        pagina = p.prompt(text='Deseja começar a partir de qual página?', default='1')
        with open(ultima_divida, 'w', encoding='utf-8') as f:
            f.write(pagina)
    
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\SIEG\\Download Divida Ativa\\ignore\\Divida Ativa",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": False,  # It will not show PDF directly in chrome
        "profile.default_content_setting_values.automatic_downloads": 1  # download multiple files
    })
    
    # iniciar o driver do chome
    while True:
        status, driver = _initialize_chrome(options)
        driver = login_sieg(driver)
        driver = sieg_iris(driver)
        driver = consulta_lista(driver, andamentos)
        if driver:
            break
        else:
            print('SIEG IRIS demorou muito pra responder, tentando novamente')
        
    driver.close()
    

if __name__ == '__main__':
    run()
