# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import os, time, re, shutil

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def consultar(options, cnpj):
    cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
    status, driver = initialize_chrome(options)
    
    # entra na página inicial da consulta
    url_inicio = 'https://www.cadesp.fazenda.sp.gov.br/(S(31mwr1eckqxyy2tfvbfuvpfc))/Pages/Cadastro/Consultas/ConsultaPublica/ConsultaPublica.aspx'
    driver.get(url_inicio)
    time.sleep(1)
    
    # clica no menu de identificação
    driver.find_element(by=By.ID, value='ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_tipoFiltroDropDownList').click()
    time.sleep(1)
    
    # clica na opção "CNPJ"
    driver.find_element(by=By.XPATH, value='/html/body/form/table[3]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/div/div[2]/div[1]/table/tbody/tr[2]/td/select/option[2]').click()
    time.sleep(1)
    
    # pega o sitekey para quebrar o captcha
    data = {'url': url_inicio, 'sitekey': '6LfWn8wZAAAAABbBsWZvt7wQXWYNTOFN3Prjcx1L'}
    response = _solve_recaptcha(data)
    
    # pega a id do campo do recaptcha
    id_response = re.compile(r'<textarea id=\"(.+)\" name=\"g-recaptcha-response')
    id_response = id_response.search(driver.page_source).group(1)
    
    # insere a solução do captcha via javascript
    driver.execute_script('document.getElementById("' + id_response + '").innerText="' + response + '"')
    
    # insere o cnpj da empresa
    _send_input('ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_valorFiltroTextBox', cnpj, driver)
    time.sleep(1)
    
    # habilita o botão de pesquisar
    driver.execute_script('document.querySelector("#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_consultaPublicaButton").removeAttribute("disabled")')
    time.sleep(1)
    
    driver.find_element(by=By.ID, value='ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_consultaPublicaButton').click()
    
    while not find_by_id('ctl00_conteudoPaginaPlaceHolder_lblCodigoControleCertidao', driver):
        time.sleep(1)
        if find_by_id('ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_MensagemErroFiltroLabel', driver):
            erro = re.compile(r'ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_MensagemErroFiltroLabel.+\">(.+)</span.')
            erro = erro.search(driver.page_source).group(1)
            driver.quit()
            print(f'Erro no login: {erro}')
            if erro == 'Captcha inválido. Tente novamente.':
                return False
            _escreve_relatorio_csv(f'{cnpj};Erro no login: {erro}', nome='Consulta ao cadastro de ICMS')
            return True
    
    print('>>> Consultando empresa')
    pega_info(cnpj, driver)
    driver.quit()
    return True


def pega_info(cnpj, driver):
    # pega ie
    try:
        ie = re.compile(r'IE:.+\n.+\"dadoDetalhe\">(.+)</td>')
        ie = ie.search(driver.page_source).group(1)
    except:
        ie = 'N/D'
    
    try:
        # pega nome
        razao = re.compile(r'Empresarial:.+\n.+\"dadoDetalhe\">(.+)</td>')
        razao = razao.search(driver.page_source).group(1)
    except:
        razao = 'N/D'
    
    try:
        # pega natureza
        natureza = re.compile(r'Jurídica:.+\n.+\"dadoDetalhe\".+>(.+)</td>')
        natureza = natureza.search(driver.page_source).group(1)
    except:
        natureza = 'N/D'
    
    try:
        # pega logradouro
        logradouro = re.compile(r'Logradouro:.+\n.+\"dadoDetalhe\".+>(.+)</td>')
        logradouro = logradouro.search(driver.page_source).group(1)
    except:
        logradouro = 'N/D'
    
    try:
        # pega número
        numero = re.compile(r'Nº:.+\n.+\"dadoDetalhe\">(.+)</td>')
        numero = numero.search(driver.page_source).group(1)
    except:
        numero = 'N/D'
    
    try:
        # pega cep
        cep = re.compile(r'CEP:.+\n.+\"dadoDetalhe\">(.+)</td>')
        cep = cep.search(driver.page_source).group(1)
    except:
        cep = 'N/D'
    
    try:
        # pega município
        municipio = re.compile(r'Município:.+\n.+\"dadoDetalhe\">(.+)</td>')
        municipio = municipio.search(driver.page_source).group(1)
    except:
        municipio = 'N/D'
    
    try:
        # pega complemento
        complemento = re.compile(r'Complemento: </b>(.+)</td>')
        complemento = complemento.search(driver.page_source).group(1)
    except:
        complemento = 'N/D'
    
    try:
        # pega bairro
        bairro = re.compile(r'Bairro:.+</b>(.+)</td>')
        bairro = bairro.search(driver.page_source).group(1)
    except:
        bairro = 'N/D'
    
    try:
        # pega UF
        uf = re.compile(r'UF:.+</b>(.+)</td>')
        uf = uf.search(driver.page_source).group(1)
    except:
        uf = 'N/D'
    
    try:
        # pega situação
        situacao = re.compile(r'Cadastral: </td>\n.+\">(.+)</td>')
        situacao = situacao.search(driver.page_source).group(1)
    except:
        situacao = 'N/D'
    
    try:
        # pega ocorrência
        ocorrencia = re.compile(r'Fiscal: .+\n.+\">(.+)</td>')
        ocorrencia = ocorrencia.search(driver.page_source).group(1)
    except:
        ocorrencia = 'N/D'
    
    try:
        # pega regime
        regime = re.compile(r'Apuração: .+\n.+\">(.+)</td>')
        regime = regime.search(driver.page_source).group(1)
    except:
        regime = 'N/D'
    
    try:
        # pega atividade
        atividade = re.compile(r'Econômicas: .+\n.+\">(.+)<br>')
        atividade = atividade.search(driver.page_source).group(1)
    except:
        atividade = 'N/D'
    
    try:
        # pega inatividade
        inatividade = re.compile(r'Inatividade:</b> (.+)</td>')
        inatividade = inatividade.search(driver.page_source).group(1)
    except:
        inatividade = 'N/D'
    
    try:
        # pega data da situação
        data_situacao = re.compile(r'Cadastral: </b>(.+)</td>')
        data_situacao = data_situacao.search(driver.page_source).group(1)
    except:
        data_situacao = 'N/D'
    
    try:
        # pega posto fiscal
        posto = re.compile(r'Fiscal:.+\">(.+)</span>')
        posto = posto.search(driver.page_source).group(1)
    except:
        posto = 'N/D'
        
    try:
        # pega Data de Credenciamento como emissor de NF-e
        credenciamento = re.compile(r'Data&nbsp;de&nbsp;Credenciamento&nbsp;como&nbsp;emissor&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>')
        credenciamento = credenciamento.search(driver.page_source).group(1)
    except:
        credenciamento = 'N/D'
        
    try:
        # pega Indicador de Obrigatoriedade de NF-e
        obrigatoriedade = re.compile(r'Indicador&nbsp;de&nbsp;Obrigatoriedade&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>')
        obrigatoriedade = obrigatoriedade.search(driver.page_source).group(1)
    except:
        obrigatoriedade = 'N/D'
        
    try:
        # pega Data de Início da Obrigatoriedade de NF-e
        inicio_obrigado = re.compile(r'Data&nbsp;de&nbsp;Início&nbsp;da&nbsp;Obrigatoriedade&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>')
        inicio_obrigado = inicio_obrigado.search(driver.page_source).group(1)
    except:
        inicio_obrigado = 'N/D'
    
    endereco = f'{logradouro}, Nº{numero}, {complemento}, {cep}, {bairro}, {municipio}-{uf}'
    _escreve_relatorio_csv(';'.join([cnpj, ie, razao.replace('&amp;', '&'), situacao, data_situacao, ocorrencia, inatividade, natureza, endereco.replace('N/D, ', '').replace(';', ''),
                                     regime, atividade.replace('<br>', ' / ').replace(';', ','), posto, credenciamento, obrigatoriedade, inicio_obrigado]), nome='Consulta ao cadastro de ICMS')
    print(f'✔ Dados coletados - {situacao} - {ocorrencia}')


@_time_execution
def run():
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
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta Débitos Municipais Jundiaí\\execucao\\Certidões",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    _escreve_header_csv(';'.join(['CNPJ', 'IE', 'NOME', 'SITUAÇÃO CADASTRAL', 'DATA DA SITUAÇÃO', 'OCORRÊNCIA FISCAL', 'DATA DE INATIVIDADE', 'NATUREZA JURÍDICA',
                                  'ENDEREÇO', 'REGIME DE APURAÇÃO', 'ATIVIDADE ECONÔMICA', 'POSTO FISCAL', 'CREDENCIAMENTO COMO EMISSOR DE NF-E', 'OBRIGATORIEDADE DE NF-E',
                                  'INÍCIO DA OBRIGATORIEDADE']), nome='Consulta ao cadastro de ICMS.csv')
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        erro = False
        while not erro:
            erro = consultar(options, cnpj)
        

if __name__ == '__main__':
    run()
