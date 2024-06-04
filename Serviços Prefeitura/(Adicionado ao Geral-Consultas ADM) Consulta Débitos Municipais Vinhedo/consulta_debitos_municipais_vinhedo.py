# -*- coding: utf-8 -*-
# from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime, shutil, os, time, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_path, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha

os.makedirs('execução/Certidões', exist_ok=True)


def pesquisar(options, cnpj, insc_muni, caminho_final):
    status, driver = _initialize_chrome(options)
    
    url_inicio = 'http://vinhedomun.presconinformatica.com.br/certidaoNegativa.jsf?faces-redirect=true'
    driver.get(url_inicio)
    
    contador = 1
    while not _find_by_id(f'homeForm:panelCaptcha:j_idt{str(contador)}', driver):
        contador += 1
        time.sleep(0.2)
        
    # resolve o captcha
    captcha = _solve_text_captcha(driver, f'homeForm:panelCaptcha:j_idt{str(contador)}')
    
    # espera o campo do tipo da pesquisa
    while not _find_by_id('homeForm:inputTipoInscricao_label', driver):
        time.sleep(1)
    # clica no menu
    driver.find_element(by=By.ID, value='homeForm:inputTipoInscricao_label').click()
    
    # espera o menu abrir
    while not _find_by_path('/html/body/div[6]/div/ul/li[2]', driver):
        time.sleep(1)
    # clica na opção "Mobiliário"
    driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/ul/li[2]').click()
    
    if not captcha:
        print('Erro Login - não encontrou captcha')

    time.sleep(1)
    # clica na barra de inscrição e insere
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').click()
    time.sleep(2)
    driver.find_element(by=By.ID, value='homeForm:inputInscricao').send_keys(insc_muni)
    _send_input(f'homeForm:panelCaptcha:j_idt{str(contador + 3)}', captcha, driver)
    time.sleep(2)
    
    # clica no botão de pesquisar
    driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[2]/div[5]/a/div').click()

    print('>>> Consultando')
    while 'dados-contribuinte-inscricao">0000000000' not in driver.page_source:
        if _find_by_path('/html/body/div[1]/div[5]/div[2]/div[1]/div/ul/li/span', driver):
            padraozinho = re.compile(r'confirmationMessage\" class=\"ui-messages ui-widget\" aria-live=\"polite\"><div '
                                     r'class=\"ui-messages-error ui-corner-all\"><span class=\"ui-messages-error-icon\"></span><ul><li><span '
                                     r'class=\"ui-messages-error-summary\">(.+).</span></li></ul></div></div><div class=\"right\"><button id=\"j_idt')
            situacao = padraozinho.search(driver.page_source).group(1)
            if situacao == 'Letras de segurança inválidas' or situacao == 'Por favor, informe o(a) Inscrição Cadastral':
                driver.close()
                return False, 'recomeçar'
            
            situacao_print = f'❌ {situacao}'
            # print(driver.page_source)
            driver.close()
            return situacao, situacao_print
        
        time.sleep(1)
    
    situacao = salvar_guia(driver, cnpj, caminho_final)
    
    situacao_print = f'✔ {situacao}'
    return situacao, situacao_print
    
    
def salvar_guia(driver, cnpj, caminho_final):
    time.sleep(1)
    while True:
        try:
            print('>>> Localizando Certidão')
            driver.find_element(by=By.XPATH, value='/html/body/div[1]/div[4]/form/div/div[4]/div[1]/a[1]/div').click()
            break
        except:
            time.sleep(1)
            pass
        
    while '<object' not in driver.page_source:
        time.sleep(1)
    
    url_pdf = re.compile(r'(/impressao\.pdf.+)\" height=\"500px\" width=\"100%\">').search(driver.page_source).group(1)
    
    print('>>> Salvando Certidão Negativa')
    driver.get('https://vinhedomun.presconinformatica.com.br/' + url_pdf)
    time.sleep(1)
    driver.close()
    
    shutil.move(os.path.join(caminho_final, 'impressao.pdf'), os.path.join(caminho_final, f'{cnpj} Certidão Negativa de Débitos Municipais Vinhedo.pdf'))
    
    return 'Certidão negativa salva'


@_time_execution
def run():
    caminho_final = "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta Débitos Municipais Vinhedo\\execução\\Certidões"
    # opções para fazer com que o chome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": caminho_final,  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)

        cnpj, insc_muni, nome= empresa
        
        while True:
            situacao, situacao_print = pesquisar(options, cnpj, insc_muni, caminho_final)
            if situacao_print != 'recomeçar':
                print(situacao_print)
                if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
                    return False
            
                _escreve_relatorio_csv(f'{cnpj};{insc_muni};{nome};{situacao}', nome='Consulta de débitos municipais de Vinhedo')
                break

    
    
if __name__ == '__main__':
    run()
