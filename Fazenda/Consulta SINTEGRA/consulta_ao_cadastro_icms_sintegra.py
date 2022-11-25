# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import os, time, re, shutil

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha


def find_by_id(xpath, driver):
    try:
        elem = driver.find_element(by=By.ID, value=xpath)
        return elem
    except:
        return None


def consultar(options, cnpj, nome):
    status, driver = _initialize_chrome(options)
    
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
    
    # clica no botão de pesquisa
    driver.find_element(by=By.ID, value='ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_consultaPublicaButton').click()
    
    # enquanto não abre a tela com as infos da empresa, verifica se aparece alguma mensagem de erro do site
    while not find_by_id('ctl00_conteudoPaginaPlaceHolder_lblCodigoControleCertidao', driver):
        time.sleep(1)
        if find_by_id('ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_MensagemErroFiltroLabel', driver):
            erro = re.compile(r'ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_MensagemErroFiltroLabel.+\">(.+)</span.')
            erro = erro.search(driver.page_source).group(1)
            driver.quit()
            print(f'Erro no login: {erro}')
            if erro == 'Captcha inválido. Tente novamente.':
                return False
            _escreve_relatorio_csv(f'{nome};{cnpj};Erro no login: {erro}', nome='Consulta ao cadastro de ICMS')
            return True
    
    # pega as infos da empresa para preencher a planilha
    print('>>> Consultando empresa')
    pega_info(cnpj, nome, driver)
    driver.quit()
    return True


def pega_info(cnpj, nome, driver):
    # pega as infos da empresa
    regex = [r'IE:.+\n.+\"dadoDetalhe\">(.+)</td>',  # inscrição estadual
             r'Empresarial:.+\n.+\"dadoDetalhe\">(.+)</td>',  # nome da empresa
             r'Cadastral: </td>\n.+\">(.+)</td>',  # situação cadastral
             r'Cadastral: </b>(.+)</td>',  # data da situação cadastral
             r'Fiscal: .+\n.+\">(.+)</td>',  # ocorrência fiscal
             r'Inatividade:</b> (.+)</td>',  # data de inatividade
             r'Jurídica:.+\n.+\"dadoDetalhe\".+>(.+)</td>',  # natureza jurídica
             r'Apuração: .+\n.+\">(.+)</td>',  # regime de apuração
             r'Econômicas: .+\n.+\">(.+)<br>',  # atividade econômica
             r'Fiscal:.+\">(.+)</span>',  # posto fiscal
             r'Data&nbsp;de&nbsp;Credenciamento&nbsp;como&nbsp;emissor&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>',  # credenciamento como emissor de nf-e
             r'Indicador&nbsp;de&nbsp;Obrigatoriedade&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>',  # obrigatoriedade de nf-e
             r'Data&nbsp;de&nbsp;Início&nbsp;da&nbsp;Obrigatoriedade&nbsp;de&nbsp;NF-e: .+\n.+\">(.+)</td>']  # início da obrigatoriedade
    infos = ''
    for reg in regex:
        try:
            info = re.compile(reg)
            info = info.search(driver.page_source).group(1).replace('&amp;', '&').replace('<br>', ' / ').replace(';', ',')
        except:
            info = 'N/D'
        infos += info + ';'
    
    # pega as infos do endereço da empresa
    regex_endereco = [r'Logradouro:.+\n.+\"dadoDetalhe\".+>(.+)</td>',  # logradouro
                      r'Nº:.+\n.+\"dadoDetalhe\">(.+)</td>',  # número
                      r'Complemento: </b>(.+)</td>',  # complemento
                      r'CEP:.+\n.+\"dadoDetalhe\">(.+)</td>',  # cep
                      r'Bairro:.+</b>(.+)</td>',  # bairro
                      r'Município:.+\n.+\"dadoDetalhe\">(.+)</td>',  # cidade
                      r'UF:.+</b>(.+)</td>']  # uf
    enderecos = ''
    for reg in regex_endereco:
        try:
            endereco = re.compile(reg)
            endereco = endereco.search(driver.page_source).group(1)
        except:
            endereco = ''
        enderecos += endereco + ', '

    enderecos = enderecos.replace(', , ', ', ').replace(';', '')
    pattern = re.compile(r'\s\s+')
    enderecos = re.sub(pattern, '', enderecos)
    
    _escreve_relatorio_csv(f"{nome};{cnpj};Ok;{infos}{enderecos}", nome='Consulta ao cadastro de ICMS')
    print(f'✔ Dados coletados')


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
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        erro = False
        while not erro:
            erro = consultar(options, cnpj, nome)
    
    _escreve_header_csv(';'.join(['CNPJ', 'IE', 'NOME', 'SITUAÇÃO CADASTRAL', 'DATA DA SITUAÇÃO', 'OCORRÊNCIA FISCAL', 'DATA DE INATIVIDADE', 'NATUREZA JURÍDICA',
                                  'REGIME DE APURAÇÃO', 'ATIVIDADE ECONÔMICA', 'POSTO FISCAL', 'CREDENCIAMENTO COMO EMISSOR DE NF-E', 'OBRIGATORIEDADE DE NF-E',
                                  'INÍCIO DA OBRIGATORIEDADE', 'ENDEREÇO']), nome='Consulta ao cadastro de ICMS.csv')


if __name__ == '__main__':
    run()
