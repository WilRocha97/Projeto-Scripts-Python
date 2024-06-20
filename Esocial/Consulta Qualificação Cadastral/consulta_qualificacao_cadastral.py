# -*- coding: utf-8 -*-
import random, time, re, os
import datetime
from selenium import webdriver
from PIL import Image
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice
from captcha_comum import _solve_recaptcha


def abre_site(options):
    # iniciar o driver do chrome
    # try:
    status, driver = _initialize_chrome(options)
    
    # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
    driver.set_page_load_timeout(60)
    
    print('>>> Acessando site')
    # aguarda o botão de habilitar a consultar aparecer
    timer = 0
    while not _find_by_id('indexForm1:botaoConsultar', driver):
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
    
    return driver, 'ok'
    

def adiciona_cadastro(driver, nome, cpf, pis, data_nasc):
    # lista de campos para preencher
    print('>>> Adicionando cadastro...')
    itens = [('formQualificacaoCadastral:nome', nome),
             ('formQualificacaoCadastral:dataNascimento', data_nasc),
             ('formQualificacaoCadastral:cpf', cpf),
             ('formQualificacaoCadastral:nis', pis)]
    
    # limpa o campo do nome para não arriscar misturar com o nome anterior caso tenha dado algum problema na hora de inseri-lo na tabela
    driver.find_element(by=By.ID, value=itens[0][0]).clear()
    time.sleep(0.1)
    
    # para cada item da lista insere a informação correspondente
    for iten in itens:
        time.sleep(random.uniform(0.5, 2))
        driver.find_element(by=By.ID, value=iten[0]).click()
        time.sleep(0.1)
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])

    # clica em adicionar cadastro a consulta
    driver.find_element(by=By.ID, value='formQualificacaoCadastral:btAdicionar').click()
    time.sleep(1)
    
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
            print('❌ O site demorou muito para responder, tentando novamente 1')
            return driver, 'erro'
        time.sleep(1)
        
    erro = re.compile(r'\"mensagem\".+class=\"erro\">(.+)</li>').search(driver.page_source)
    if erro:
        print(f'❌ {erro.group(1)}')
        return driver, erro.group(1)
    time.sleep(1)
    timer += 1
    # se passar 60 segundos e não aparecer a tabela, retorna um erro
    if timer > 60:
        print('❌ O site demorou muito para responder, tentando novamente 1')
        return driver, 'erro'
    
    print('>>> Cadastro adicionado')
    return driver, 'ok'
    
    
def validar_consulta(driver):
    print('\n>>> Validando dados...')
    # clica em validar a consulta
    driver.find_element(by=By.ID, value='formValidacao2:botaoValidar2').click()
    timer = 0

    while not _find_by_id('formValidacao:botaoValidar', driver):
        time.sleep(1)
        timer += 1
        # se passar 60 segundos e não carregar o site, retorna um erro
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente 2')
            return driver, 'erro'
    
    # resolver o captcha
    data = {'url': 'https://consultacadastral.inss.gov.br/Esocial/pages/qualificacao/qualificar.xhtml',
            'sitekey': '6LeHKtspAAAAAIGfAB4kjt6ApvYLB8w-lSRlomzg'}
    captcha = _solve_recaptcha(data)
    
    if not captcha:
        print('❌ Erro Login - não encontrou captcha')
        return driver, 'erro captcha'

    # insere a resposta do captcha e clica em validar
    #try:
    #driver.find_element(by=By.ID, value='g-recaptcha-response').send_keys(captcha)
    driver.execute_script('document.getElementById("g-recaptcha-response").innerText="' + captcha + '"')
    driver.find_element(by=By.ID, value='formValidacao:botaoValidar').click()
    """except:
        print('❌ Erro ao validar a consulta, tentando novamente')
        return driver, 'erro'"""

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
            print('❌ O site demorou muito para responder, tentando novamente 3')
            return driver, 'erro'
    return driver, 'ok'


def consulta(driver, cadastros_pesquisados):
    print('>>>Analisando dados...')
    page_source_filtrado = ''
    
    # modifica o código do site para que poder criar um regex mais eficiente
    page_source_modificado = (driver.page_source.replace('<br>', '').replace('class="tamanho', 'class="tamanho\n').replace('<', '<\n'))
    for linha in page_source_modificado.split('\n'):
        if re.compile(r'\d\d\">(.+)').search(linha):
            page_source_filtrado = page_source_filtrado + linha + '\n'
    
    page_source_filtrado = page_source_filtrado.replace('/span><\n', '').replace('/td><\n', '').replace('span class="tamanho\n', '').replace('td class="left"><\n', '').replace('td class="center"><\n', '')
    
    #print(page_source_filtrado)
    #time.sleep(22)
    
    # o seguinte método não é necessário, mas preferí ter uma segurança maior nos dados pesquisados
    # adiciona as informações capturadas do site em um dicionário para que poder serem buscadas seguindo os CPFs usados para preencher os dados da consulta
    cadastros = {}
    resultados = (re.compile(r'20\">(.+)<\n\d\d\">(\d\d/\d\d/\d\d\d\d)<\n\d\d\">(\d\d\d.\d\d\d.\d\d\d-\d\d)<\n\d\d\">(\d.\d\d\d.\d\d\d.\d\d\d-\d)<\n\d\d\"(.+)<\n\d\d\"(.+)<\n\d\d\"(.+)<\n\d\d\"(.+)<')
                  .findall(page_source_filtrado))
    for resultado in resultados:
        cadastros[str(resultado[2]).replace('-', '').replace('.', '')] = (
            resultado[0],
            resultado[1],
            resultado[3],
            f'{resultado[4].replace(">", "")} {resultado[5].replace(">", "")}',
            f'{resultado[6].replace(">", "")} {resultado[7].replace(">", "")}')
    
    # pesquisa os CPFs da lista de consulta no dicionário das respostas para garantir que foram todos pesquisados
    for cadastro_pesquisado in cadastros_pesquisados:
        print(cadastro_pesquisado[0], f'{cadastros[cadastro_pesquisado[0]][3].strip()}')
        _escreve_relatorio_csv(f'{cadastro_pesquisado[0]};{cadastro_pesquisado[1]};{cadastro_pesquisado[2]};{cadastro_pesquisado[3]};{cadastro_pesquisado[4]};{cadastro_pesquisado[5]};'
                               f'{cadastros[cadastro_pesquisado[0]][3].strip()};{cadastros[cadastro_pesquisado[0]][4]}', nome='Consulta Qualificação Cadastral')
    
    print('\n')
    return driver


def verifica_dados(cpf, nome, cod_empresa, cod_empregado, pis, data_nasc):
    infos = [(cpf, 'CPF'),
             (nome, 'Nome'),
             (pis, 'PIS'),
             (data_nasc, 'Data de nascimento')]

    for info in infos:
        if info[0] == '':
            print(f'❌ {info[1]} não informado')
            _escreve_relatorio_csv(f'{cpf};{nome};{cod_empresa};{cod_empregado};{pis};{data_nasc};{info[1]} não informado', nome='Consulta Qualificação Cadastral')
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
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    
    contador = 1
    cadastros_pesquisados = []
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        cpf, nome, cod_empresa, cod_empregado, pis, data_nasc = empresa
        # verifica se não tem nenhum dado faltando
        if not verifica_dados(cpf, nome, cod_empresa, cod_empregado, pis, data_nasc):
            continue
            
        if contador == 1:
            driver, resultado = abre_site(options)
        
        time.sleep(random.uniform(0.5, 2))
        
        # faz login no site
        driver, resultado = adiciona_cadastro(driver, nome, cpf, pis, data_nasc)
        if resultado != 'ok':
            _escreve_relatorio_csv(f'{cpf};{nome};{cod_empresa};{cod_empregado};{pis};{data_nasc};{resultado}', nome='Consulta Qualificação Cadastral')
            continue
        
        # adiciona em uma lista os dados inseridos na consulta para verificar se os CPFs conferem com a resposta do site
        #cpf_pontuado = re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', cpf)
        cadastros_pesquisados.append((cpf, nome, cod_empresa, cod_empregado, pis, data_nasc))
        
        contador += 1
        # a cada 10 CPFs, realiza a consulta, agora o site aceita até 10 cadastros por vês, porém, colocando uma espera aleatória no preenchimento dos dados para que o site não bloqueie como acesso suspeito.
        if contador > 10:
            driver, resultado = validar_consulta(driver)
            driver = consulta(driver, cadastros_pesquisados)
            driver.close()
            
            contador = 1
            cadastros_pesquisados = []
    
    driver, resultado = validar_consulta(driver)
    driver = consulta(driver, cadastros_pesquisados)
    driver.close()
    
    _escreve_header_csv('CPF;NOME;CÓD EMPRESA;CÓD EMPREGADO;PIS;DATA DE NASCIMENTO;SITUAÇÃO;OBSERVAÇÕES', nome='Consulta Qualificação Cadastral')


if __name__ == '__main__':
    run()
