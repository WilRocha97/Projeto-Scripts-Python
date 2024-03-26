import time, re, os
from bs4 import BeautifulSoup
from xhtml2pdf import pisa
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from pyautogui import alert

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _get_comp
from chrome_comum import _find_by_id, _find_by_path
from fazenda_comum import _new_session_fazenda_driver
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def captura_dados(driver):
    print('>>> Capturando dados')
    # define os regex dos campos únicos na consulta
    if re.compile(r'Responsável pela Entrega:</span>\s+CPF:.+\s+(.+)').search(driver.page_source):
        regex_responsavel = r'Responsável pela Entrega:</span>\s+CPF:.+\s+(.+)'
    else:
        regex_responsavel = r'Responsável pela Entrega:</span>\s+(.+).+/td>'
    
    if re.compile(r"Categoria:</span>\s+()</td>").search(driver.page_source):
        regex_categoria = r'Categoria:</span>\s+()</td>'
    else:
        regex_categoria = r'Categoria:</span>\s+(.+).+/td>'
        
    if re.compile(r'(O contribuinte está dispensado da entrega da GIA nessa referência)').search(driver.page_source):
        regex_chave = r'(O contribuinte está dispensado da entrega da GIA nessa referência)'
    else:
        regex_chave = r'Chave de Autenticação:</span>\s+(.+).+/td>'
    
    # lista para concatenar os itens da consulta em uma única variável
    termos = [r'Razão Social:</span>\s+(.+).+/td>',
              r'Data de Entrega:</span>\s+(.+).+/td>',
              regex_responsavel,
              regex_chave,
              r'CNAE:</span>\s+(.+).+/td>',
              r'Origem:</span>\s+(.+).+/td>',
              r'Referência:</span>\s+(.+).+/td>',
              r'Tipo:</span>\s+(.+)',
              regex_categoria,
              r'Protocolo:</span>\s+.+\">(.+).+/a>']
    
    resultado = ''
    # tenta procurar e concatenar os itens, se não conseguir printa o código da página para análise
    try:
        for termo in termos:
            resposta = re.compile(termo).search(driver.page_source).group(1)
            if resposta == '':
                resposta = 'Não consta'
                
            resultado += resposta + ';'
    except:
        print(driver.page_source)
        time.sleep(22)
        
    return driver, resultado.replace('&amp;', '&')


def create_pdf(driver, nome_arquivo, comp_formatado):
    # salva o pdf criando ele a partir do código html da página, para que o PDF criado seja editável
    e_dir_pdf = os.path.join('execução', 'Arquivos ' + comp_formatado)
    os.makedirs(e_dir_pdf, exist_ok=True)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    # Remove atributo href das tags, links e urls que não serão usados
    for tag in soup.find_all(href=True):
        tag['href'] = ''

    # Remove tags img e script, não serão usados
    _ = list(tag.extract() for tag in soup.find_all('img'))
    _ = list(tag.extract() for tag in soup.find_all('script'))
    
    # remove mais algumas coisas específicas
    soup = str(soup)\
        .replace('target="_blank">Cidadão SP</a></td>', '')\
        .replace('target="_blank">saopaulo.sp.gov.br</a></td>', '') \
        .replace('<a class="botao-sistema" href="">Home</a>', '') \
        .replace('<a class="botao-sistema" href="">Imprimir</a>', '') \
        .replace('<a class="botao-sistema" href="">Encerrar</a>', '')\
        .replace('target="_blank">Ouvidoria</a></td>', '')\
        .replace('target="_blank">Transparência</a></td>', '')\
        .replace('href="" target="_blank">SIC</a></td>', '')
    
    soup = BeautifulSoup(soup, 'html.parser')
    
    with open(os.path.join(e_dir_pdf, nome_arquivo), 'w+b') as pdf:
        pisa.showLogging()
        pisa.CreatePDF(str(soup), pdf)
    
    return driver, 'Arquivo gerado'


def consulta_gia(ie, comp, driver, sid):
    print(f'>>> Consultando entrega da GIA')
    comp = comp.split('/')
    mes = comp[0]
    if mes[0] == '0': mes = mes[1]
    ano = comp[1]
    
    driver.get('https://cert01.fazenda.sp.gov.br/novaGiaWEB/consultaIe.gia?method=init&SID=' + sid + '&servico=GIA')
    
    while not _find_by_id('ie', driver):
        time.sleep(1)
    
    driver.find_element(by=By.ID, value='ie').send_keys(ie)
    
    # lista com o xpath do elemento e a informação que ele deve receber
    drops = [('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/select[1]', mes),
             ('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/select[1]', mes),
             ('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/select[2]', ano),
             ('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[3]/td[2]/select[2]', ano),
             ]
    
    for drop in drops:
        # Localize o elemento do dropdown
        dropdown_element = driver.find_element(by=By.XPATH, value=drop[0])
        # Crie um objeto Select com o elemento do dropdown
        dropdown = Select(dropdown_element)
        # Selecione o item desejado pelo valor
        dropdown.select_by_value(drop[1])
    
    # clica para confirmar a consulta
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='/html/body/form/table/tbody/tr[6]/td/input').click()
    
    time.sleep(1)
    # captura erros do site, se achar retorna a mensagem
    resultado = re.compile(r'MENSAGEM DO SISTEMA: (.+)').search(driver.page_source)
    if not resultado:
        resultado = re.compile(r'RESULTADO-ERRO\">(.+)</td>').search(driver.page_source)
        if not resultado:
            resultado = re.compile(r'RESULTADO-ERRO.+\n(.+)').search(driver.page_source)
            if not resultado:
                return driver, 'ok'
            
    return driver, resultado.group(1)
        
    
@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    
    comp = _get_comp(printable='mm/yyyy', strptime='%m/%Y')
    comp_formatado = comp.replace("/", "-")
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o índice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    sid = ''
    driver = False
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, ie, usuario, senha, perfil = empresa
        
        # printa o índice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        
        # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
        if usuario_anterior != usuario:
            # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
            if driver:
                driver.close()
                
            contador = 0
            # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
            while True:
                try:
                    # cookies, sid = _new_session_fazenda_driver(usuario, senha, perfil)
                    driver, sid = _new_session_fazenda_driver(usuario, senha, perfil, retorna_driver=True, options=options)
                    if driver == 'erro':
                        break
                    break
                except:
                    print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
                
                time.sleep(1)
            
            # se der erro no login
            if driver == 'erro':
                # se for contador exibe um alerta e para a execução
                if perfil == 'contador':
                    alert(f'Contador!: {sid}')
                    return
                # se for contribuinte so ano na planilha
                else:
                    _escreve_relatorio_csv(f"{cnpj};{ie};{sid.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')}", nome=f'Consulta Envio de GIA - {comp_formatado}')
                    continue
                    
        driver, resultado = consulta_gia(ie, comp, driver, sid)
        
        # se a consulta retornar um ok, cria o PDF e captura os dados
        if resultado == 'ok':
            nome_arquivo = f'{cnpj} - Entrega de GIA {comp_formatado}.pdf'
            driver, resultado = create_pdf(driver, nome_arquivo, comp_formatado)
            driver, resultado = captura_dados(driver)
        else:
            _escreve_relatorio_csv(f'{cnpj};{ie};{usuario};{senha};{perfil}', nome=f'Próximos dados', local='ignore')
        
        # escreve na planilha de andamentos o resultado da execução atual
        try:
            _escreve_relatorio_csv(f"{cnpj};{ie};{resultado.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')}",
                                   nome=f'Consulta Envio de GIA - {comp_formatado}')
        except:
            raise Exception(f"Erro ao escrever esse texto: {resultado}")
        print(resultado)
        usuario_anterior = usuario

    driver.close()
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;IE;NOME;DATA ENTREGA;RESPONSÁVEL;CHAVE DE AUTENTICAÇÃO;CNAE;ORIGEM;REFERÊNCIA;TIPO;CATEGORIA;PROTOCOLO',
                        nome=f'Consulta Envio de GIA - {comp_formatado}')
    return True


if __name__ == '__main__':
    run()
