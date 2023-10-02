# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from xhtml2pdf import pisa
from bs4 import BeautifulSoup
import os, time, re, fitz, shutil

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_text_captcha
from chrome_comum import _initialize_chrome, _send_input, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, cnpj):
    print('>>> Logando no site')
    # timer de espera do captcha com tentativas de carregar o site
    # se o site der erro ou demorar mais de 60 segundos para aparecer o captcha retorna o erro e tenta novamente
    timer = 0
    while not _find_by_id('captcha_image', driver):
        try:
            driver.get('https://situacao.campinas.sp.gov.br/')
        except:
            print('>>> Site demorou pra responder, tentando novamente')
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'
    
    # resolver o captcha
    captcha = _solve_text_captcha(driver, 'captcha_image')
    
    # clica para habilitar consulta por CNPJ
    driver.find_element(by=By.ID, value='radio3').click()
    # insere o CNPJ
    _send_input('xInscricao', cnpj, driver)
    # insere a resposta do captcha
    _send_input('cap_text', captcha, driver)
    # clica para consultar
    driver.find_element(by=By.XPATH, value='/html/body/form/table[2]/tbody/tr[3]/td/table[2]/tbody/tr[1]/td[1]/input').click()
    
    # enquanto a tabela com as informações do cadastro não abre, verifica mensagens de erro
    while not _find_by_path('/html/body/div/table/tbody/tr/td[2]', driver):
        # verifica se a consulta foi realizada
        if re.compile(r'Consulta não realizada').search(driver.page_source):
            causas = re.compile(r'(Possíveis causas.+)<br></font></td>').search(driver.page_source).group(1)
            return driver, 'sem cadastro', f'❗ Consulta não realizada. {causas}'
        
        # verifica se existe mensagem de erro do captcha
        if re.compile(r'Número digitado não corresponde ao Número de Acesso.').search(driver.page_source):
            print('>>> Erro no captcha, tentando novamente')
            return driver, 'erro captcha', ''
        
        if re.compile(r'id=(.+)\">Visualizar').findall(driver.page_source):
            return driver, 'ok', ''
            
        time.sleep(1)
    
    print('>>> Cadastro encontrado')
    return driver, 'ok', ''


def create_pdf(driver, download_folder, cnpj):
    print('>>> Criando PDF do comprovante')
    # pega a inscrição cadastral
    inscricao_cadastro_site = re.compile(r'INSCRIÇÃO MOBILIÁRIA MUNICIPAL.+\s\s(\d.+) </td>').search(driver.page_source).group(1)
    #cria o nome do arquivo
    nome_arquivo = f'Comprovante de Situação Cadastral no Município de Campinas - {cnpj} - {inscricao_cadastro_site}.pdf'
    # salva o pdf criando ele a partir do código html da página, para que o PDF criado seja editável
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    imagens = soup.findAll('img')
    for img in imagens:
        img.extract()
        
    with open(os.path.join(download_folder, nome_arquivo), 'w+b') as pdf:
        for s in soup:
            pisa.showLogging()
            pisa.CreatePDF(str(s), pdf)
    
    return driver, '✔ Comprovante salvo', nome_arquivo


def coleta_dados_comprovante(download_folder, nome_arquivo):
    print('>>> Analisando PDF')
    # abre o PDF da nota para capturar algumas infos para adicionar na planilha de andamentos e para renomear o arquivo
    with fitz.open(os.path.join(download_folder, nome_arquivo)) as pdf:
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            
            dados = [r'INSCRIÇÃO MOBILIÁRIA MUNICIPAL\n(.+)',
                     r'SUBSTITUTO TRIBUTÁRIO DE TODOS OS SERVIÇOS TOMADOS\n(.+)',
                     r'SITUAÇÃO CADASTRAL\n(.+)',
                     r'DATA DE INÍCIO DAS ATIVIDADES\n(\d\d/\d\d/\d\d\d\d)',
                     r'DATA DE ENCERRAMENTO\n(\d\d/\d\d/\d\d\d\d)',
                     r'DATA DA ÚLTIMA ALTERAÇÃO\n(\d\d/\d\d/\d\d\d\d)']
            
            dados_arquivo = ''
            for dado in dados:
                try:
                    dados_arquivo += re.compile(dado).search(textinho).group(1) + ';'
                except:
                    dados_arquivo += ';'
                    
    return dados_arquivo


@_time_execution
def run():
    download_folder = "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta de Situação Cadastral Mobiliária Campinas\\execução\\Comprovantes"
    os.makedirs(download_folder, exist_ok=True)
    
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        resultado_consulta = ''
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 15 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            
            driver, situacao, mensagem = login(driver, cnpj)
            if situacao == 'sem cadastro':
                resultado_consulta = f'{cnpj};{nome};{mensagem[2:]}'
                break
            
            # se o login estiver ok, cria o PDF do comprovante, pois logo após o login, ele já é exibido
            # depois coleta os dados necessários para escrever na planilha
            if situacao == 'ok':
                # se retornar ok, verifica se existe mais de um cadastro e consulta cada um
                cadastros = re.compile(r'id=(.+)\">Visualizar').findall(driver.page_source)
                # se encontrar mais de um cadastro
                if cadastros:
                    # para cada cadastro encontrado
                    for count_2, id_cadastro in enumerate(cadastros):
                        final = '\n'
                        if count_2 == len(cadastros) - 1:
                            final = ''
                            
                        # clica no botão de visualizar
                        driver.get('https://situacao.campinas.sp.gov.br/visualizar.php?id=' + id_cadastro)
                        # espera a tabela abrir
                        while not _find_by_path('/html/body/div/table/tbody/tr/td[2]', driver):
                            time.sleep(1)
                            
                        driver, mensagem, nome_arquivo = create_pdf(driver, download_folder, cnpj)
                        dados = coleta_dados_comprovante(download_folder, nome_arquivo)
                        resultado_consulta += f'{cnpj};{nome};{mensagem[2:]};{dados}{final}'
                    break
                    
                # se não tiver lista de cadastro
                else:
                    driver, mensagem, nome_arquivo = create_pdf(driver, download_folder, cnpj)
                    dados = coleta_dados_comprovante(download_folder, nome_arquivo)
                    resultado_consulta = f'{cnpj};{nome};{mensagem[2:]};{dados}'
                    break
                
            # fecha o driver dentro do while para tentar novamente com outro driver
            driver.close()
        
        # toda escrita na planilha e print com os resultados será feito aqui, todas as funções retornaram esses dados
        # apenas prints de referência serão feitos nas funções
        print(mensagem)
        _escreve_relatorio_csv(resultado_consulta)
        driver.close()
    
    # escreve o cabeçalho da planilha
    _escreve_header_csv('CNPJ;'
                        'NOME;'
                        'RESULTADO;'
                        'SUBSTITUTO TRIBUTÁRIO DE TODOS OS SERVIÇOS TOMADOS;'
                        'SITUAÇÃO CADASTRAL;'
                        'DATA DE INÍCIO DAS ATIVIDADES;'
                        'DATA DE ENCERRAMENTO;'
                        'DATA DA ÚLTIMA ALTERAÇÃO')
    

if __name__ == '__main__':
    run()
