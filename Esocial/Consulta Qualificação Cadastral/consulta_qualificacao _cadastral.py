# -*- coding: utf-8 -*-
import time, re, os
from selenium import webdriver
from PIL import Image
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_text_captcha


def login(driver, nome, cpf, pis, data_nasc):
    print('>>> Acessando site')
    # aguarda o botão de habilitar a consultar aparecer
    timer = 0
    while not _find_by_id('indexForm1:botaoConsultar', driver):
        print('🕤')
        # abre o site da consulta e caso de erro é porque o site demorou pra responder,
        # nesse caso retorna um erro para tentar novamente
        try:
            driver.get('http://consultacadastral.inss.gov.br/Esocial/pages/index.xhtml')
        except:
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'
        
    # clica para habilitar a consulta
    driver.find_element(by=By.ID, value='indexForm1:botaoConsultar').click()

    # aguarda o campo de nome aparecer
    timer = 0
    while not _find_by_id('formQualificacaoCadastral:nome', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    # lista de campos para preencher
    itens = [('formQualificacaoCadastral:nome', nome),
             ('formQualificacaoCadastral:dataNascimento', data_nasc),
             ('formQualificacaoCadastral:cpf', cpf),
             ('formQualificacaoCadastral:nis', pis)]

    # para cada item da lista insere a informação correspondente
    for iten in itens:
        driver.find_element(by=By.ID, value=iten[0]).click()
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])

    # clica em adicionar cadastro a consulta
    # não é possível adicionar mais cadastros a pesquisa, pois o site é um lixo e sai do ar
    driver.find_element(by=By.ID, value='formQualificacaoCadastral:btAdicionar').click()
    print('>>> Acessando cadastro')
    timer = 0
    # espera a tabela com as infos do funcionário aparecer, se aparecer uma mensagem, pega o conteúdo dela e retorna
    while not _find_by_id('gridDadosTrabalhador', driver):
        erro = re.compile(r'\"mensagem\".+class=\"erro\">(.+)</li>').search(driver.page_source)
        if erro:
            print(f'❌ {erro.group(1)}')
            return driver, erro.group(1)
        
        time.sleep(1)
        timer += 1
        # se passar 60 segundos e não aparecer a tabela, retorna um erro
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente')
            return driver, 'erro'

    # clica em validar a consulta
    driver.find_element(by=By.ID, value='formValidacao2:botaoValidar2').click()
    
    timer = 0
    # tira um print da tela com e recorta apenas a imagem do captcha para enviar para a api
    while not _find_by_id('captcha_challenge', driver):
        time.sleep(1)
        timer += 1
        # se passar 60 segundos e não carregar o site, retorna um erro
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente')
            return driver, 'erro'
    
    # resolver o captcha
    captcha = _solve_text_captcha(driver, 'captcha_challenge')
    
    if not captcha:
        print('❌ Erro Login - não encontrou captcha')
        return driver, 'erro captcha'

    # insere a resposta do captcha e clica em validar
    try:
        driver.find_element(by=By.ID, value='captcha_campo_resposta').send_keys(captcha)
        driver.find_element(by=By.ID, value='formValidacao:botaoValidar').click()
    except:
        print('❌ Erro ao validar a consulta, tentando novamente')
        return driver, 'erro'

    print('>>> Acessando cadastro')
    # aguarda as informações do cadastro aparecerem, se demorar mais de 1 minuto ou a resposta do captcha estiver errada,
    # retorna um erro e tenta novamente
    timer = 0
    while not _find_by_id('j_idt21:gridDadosTrabalhador', driver):
        time.sleep(1)
        timer += 1
        if re.compile(r'captcha: A resposta .+ não corresponde ao desafio gerado').search(driver.page_source):
            print('❌ A resposta não corresponde ao desafio gerado, tentando novamente')
            return driver, 'erro'
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente')
            return driver, 'erro'

    return driver, 'ok'


def consulta(driver):
    # captura a situação cadastral do funcionário pesquisado
    try:
        mensagem = re.compile(r'</span><span class=\"tamanho\d+.+>(.+)<br><br></span></td><td').search(driver.page_source).group(1)
        return driver, str(mensagem).replace(" : ", ": ")
    except:
        pass

    mensagens_regex = [r'</span><span class=\"tamanho\d+.+>(.+)<br><br></span></td><td class=\"left\"><span class=\"tamanho\d+\"></span><span class=\"tamanho\d+\">(.+)<br><br> </span></td></tr>',
                       r'<span class=\"tamanho\d+.+>(.+)<br><br></span><span class=\"tamanho\d+\"></span></td><td class=\"left\"><span class=\"tamanho\d+\">(.+)<br><br></span><span class=\"tamanho\d+\"> </span></td></tr>']

    for mensagem_regex in mensagens_regex:
        try:
            mensagens = re.compile(mensagem_regex).search(driver.page_source)
            mensagem = f'{mensagens.group(1)};{mensagens.group(2).replace(":<br>", " | ").replace("<br>", " | ").replace(" |  | ", " | ")}'
            return driver, str(mensagem)
        except:
            pass

    print(driver.page_source)
    return driver, str('Erro ao analisar o cadastro')


def verifica_dados(cpf, nome, cod_empresa, cod_empregado, pis, data_nasc):
    infos = [(cpf, 'CPF'),
             (nome, 'Nome'),
             (pis, 'PIS'),
             (data_nasc, 'Data de nascimento')]

    for info in infos:
        if info[0] == '':
            print(f'❌ {info[1]} não informado')
            _escreve_relatorio_csv(f'{cpf};{nome};{cod_empresa};{cod_empregado};{pis};{data_nasc};{info[1]} não informado')
            return False

    return True


@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1366,768')
    # options.add_argument("--start-maximized")

    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa, index)
        cpf, nome, cod_empresa, cod_empregado, pis, data_nasc = empresa
        
        # verifica se não tem nenhum dado faltando
        if not verifica_dados(cpf, nome, cod_empresa, cod_empregado, pis, data_nasc):
            continue

        while True:
            # iniciar o driver do chrome
            status, driver = _initialize_chrome(options)
            
            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(60)
            
            # faz login no site
            driver, resultado = login(driver, nome, cpf, pis, data_nasc)
            # se não der erro no login, sai do while e realiza a consulta
            if resultado != 'erro':
                break
            driver.close()
        
        if resultado == 'ok':
            driver, resultado = consulta(driver)
            
        print(f'❕ {resultado}')
        _escreve_relatorio_csv(f'{cpf};{nome};{cod_empresa};{cod_empregado};{pis};{data_nasc};{resultado}')
        driver.close()

    _escreve_header_csv('CPF;NOME;CÓD EMPRESA;CÓD EMPREGADO;PIS;DATA DE NASCIMENTO;SITUAÇÃO;OBSERVAÇÕES')


if __name__ == '__main__':
    run()
