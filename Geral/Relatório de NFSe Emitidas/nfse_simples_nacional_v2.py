# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _send_input, _find_by_id
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def login(driver, usuario, senha):
    print('>>> Logando no site')
    try:
        driver.get('https://www.nfse.gov.br/EmissorNacional/Login?ReturnUrl=%2fEmissorNacional')
    except:
        print('>>> Site demorou pra responder, tentando novamente')
        return driver, 'erro'
    time.sleep(1)
    
    # insere o usuário e a senha no site
    _send_input('Inscricao', usuario, driver)
    _send_input('Senha', senha, driver)
    time.sleep(1)
    
    # clica no botão para logar
    driver.find_element(by=By.XPATH, value='/html/body/section/div/div/div[2]/div[2]/div[1]/div/form/div[3]/button').click()
    time.sleep(1)
    return driver, 'ok'

    
def consulta_notas(cnpj, nome, driver):
    print('>>> Consultando notas emitidas')
    driver.get('https://www.nfse.gov.br/EmissorNacional/Notas/Emitidas')
    time.sleep(1)
    
    link_notas = re.compile(r'href=\"(/EmissorNacional/Notas/Visualizar/Index/.+)\" class=\"list-group-item\">').findall(driver.page_source)
    
    if not link_notas:
        print(driver.page_source)
    
    for count, link_nota in enumerate(link_notas, start=1):
        print(f'>>> Abrindo {count}° nota')
        driver.get('https://www.nfse.gov.br' + link_nota)
        time.sleep(1)
        
        data_de_emissao = re.compile(r'Data de emissão</span></label><span class=\"form-control-static texto\">(.+) \n\s+(.+)\n\s+(.+)</span></div>').search(driver.page_source)
        emissao_nota = f'{data_de_emissao.group(1)} {data_de_emissao.group(2)} {data_de_emissao.group(3)}'
        
        dados_nota = [r'Chave de acesso</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Data de emissão</span></label><span class=\"form-control-static texto\">(.+) \n\s+(.+)\n\s+(.+)</span></div>',
                      r'Versão</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Número</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Série</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Emitente.+(\n.+){4}Razão Social</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Emitente.+(\n.+){9}CNPJ</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Emitente.+(\n.+){12}Inscrição Municipal</span></label><span class=\"form-control-static .+\">\n\s+(.+)',
                      r'Situação Perante o Simples Nacional</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Regime Especial de Tributação</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Endereço do Estabelecimento/Domicílio</span></label><span class=\"form-control-static .+\">(.+\n.+\n.+\n.+\n.+\n.+)</span></div>',
                      r'Telefone</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Emitente.+(\n.+){41}Email</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Tributação do ISSQN</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Município de Incidência</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'País Resultado da Prestação de Serviço</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'Tipo de Imunidade</span></label><span class=\"form-control-static .+\">(.+)</span></div>',
                      r'']
        
        print(driver.page_source)
        time.sleep(55)
        _escreve_relatorio_csv(dados)
        print(dados.replace(';', ' - '))
        
    return driver


@_time_execution
def run():
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, usuario, senha = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        while True:
            status, driver = _initialize_chrome(options)
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(15)
            driver, situacao = login(driver, usuario, senha)
            if situacao == 'ok':
                driver = consulta_notas(cnpj, nome, driver)
                break
            driver.close()
        driver.close()
    
    _escreve_header_csv('CHAVE DE ACESSO;DATA DE EMISSÃO;VERSÃO;NÚMERO NFE;SÉRIE;RAZÃO SOCIAL;CNPJ;SITUAÇÃO SIMPLES NACIONAL;'
                        'REGIME ESPECIAL DE TRIBUTAÇÃO;TELEFONE;EMAIL;TRIBUTAÇÃO DE ISSQN;PAÍS DA PRESTAÇÃO DE SERVIÇO;MUNICÍPIO DE INCIDÊNCIA;'
                        'TIPO DE IMUNIDADE;SUSPENSÃO DE ISSQN;NÚMERO PROCESSO SUSPENSÃO;BENEFÍCIO MUNICIPAL;VALOR DO SERVIÇO;TOTAL DEDUÇÕES/REDUÇÕES;'
                        'TOTAL BENEFÍCIO MUNICIPAL;RETENÇÃO;VERSÃO DA APLICAÇÃO;AMBIENTE GERADOR;SITUAÇÃO DA NFS-E;CNPJ TOMADOR;RAZÃO SOCIAL TOMADOR;'
                        'EMAIL TOMADOR;PAÍS DO SERVIÇO;MUNICÍPIO DO SERVIÇO;CÓDIGO DE TRIBUTAÇÃO NACIONAL;ITEM DA NBS CORRESPONDENTE;'
                        'DESCRIÇÃO DO SERVIÇO;''NOME DO EVENTO;DATA DE INÍCIO;DATA DO FIM;CÓDIGO DE IDENTIFICAÇÃO;'
                        'Nº DOCUMENTO DE RESPONSABILIDADE TÉCNICA;DOCUMENTO DE REFERÊNCIA;INFOS COMPLEMENTARES;SITUAÇÃO TRIBUTÁRIA DO PIS/COFINS;'
                        'TIPO DE RETENÇÃO DO PIS/COFINS;BASE DE CÁLCULO DO PIS/COFINS;PIS ALÍQUOTA;PIS VALOR DO IMPOSTO;COFINS ALÍQUOTA;'
                        'COFINS VALOR DO IMPOSTO;VALOR RETIDO IRPF;VALOR RETIDO CSLL;VALOR RETIDO CP;OPÇÃO TOTAL DOS TRIBUTOS')
        
if __name__ == '__main__':
    run()
