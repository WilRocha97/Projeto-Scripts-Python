import time, re, os
from sys import path
from time import sleep
from PIL import Image
from fpdf import FPDF

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

path.append(r'..\..\_comum')
from pyautogui_comum import _get_comp
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from fazenda_comum import _new_session_fazenda_driver
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice


def captura_dados(driver):
    print('>>> Capturando dados')
    # coleta todos os dados padronizados nos dados da página
    if re.compile(r'Responsável pela Entrega:</span>\s+CPF:.+\s+(.+)').search(driver.page_source):
        regex_responsavel = r'Responsável pela Entrega:</span>\s+CPF:.+\s+(.+)'
    else:
        regex_responsavel = r'Responsável pela Entrega:</span>\s+(.+).+/td>'
    
    if re.compile(r'Categoria:</span>\s+</td>').search(driver.page_source):
        regex_categoria = r'Categoria:</span>\s+</td>'
    else:
        regex_categoria = r'Categoria:</span>\s+(.+).+/td>'
        
    if re.compile(r'(O contribuinte está dispensado da entrega da GIA nessa referência)').search(driver.page_source):
        regex_chave = r'(O contribuinte está dispensado da entrega da GIA nessa referência)'
    else:
        regex_chave = r'Chave de Autenticação:</span>\s+(.+).+/td>'
        
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
    try:
        for termo in termos:
            resultado += re.compile(termo).search(driver.page_source).group(1) + ';'
    except:
        print(driver.page_source)
        time.sleep(22)
        
    return driver, resultado.replace('&amp;', '&')


def create_pdf(driver, nome_arquivo):
    print('>>> Salvando PDF')
    e_dir_print = os.path.join('ignore', 'print consulta')
    e_dir_pdf = os.path.join('execução', 'Arquivos')
    os.makedirs(e_dir_print, exist_ok=True)
    os.makedirs(e_dir_pdf, exist_ok=True)
    
    # Capturar a captura de tela da página
    screenshot_file = os.path.join(e_dir_print, 'screenshot.png')
    driver.save_screenshot(screenshot_file)
    
    # Converter a captura de tela em PDF
    pdf_file = os.path.join(e_dir_pdf, nome_arquivo + '.pdf')
    pdf = FPDF()
    pdf.add_page()
    pdf.image(screenshot_file, 0, 0, pdf.w, pdf.h)
    pdf.output(pdf_file, 'F')
    
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
    time.sleep(2)
    driver.find_element(by=By.XPATH, value='/html/body/form/table/tbody/tr[6]/td/input').click()
    
    time.sleep(1)
    # captura mensagens de erro, se encontrar retorna a mensagem se não retorna um ok
    try:
        # Aguardar o alerta ser exibido
        alert = driver.switch_to.alert
        
        # Obter a mensagem do alerta
        mensagem = alert.text
        
        # Imprimir a mensagem do alerta
        print(f'❌ {mensagem}')
        
        # Aceitar o alerta
        alert.accept()
        return driver, mensagem
    except:
        resultado = re.compile(r'MENSAGEM DO SISTEMA: (.+)').search(driver.page_source)
        if not resultado:
            resultado = re.compile(r'RESULTADO-ERRO.+\n(.+)').search(driver.page_source)
            if not resultado:
                return driver, 'ok'
            
        return driver, resultado.group(1)
        
    
@_time_execution
def run():
    comp = _get_comp(printable='mm/yyyy', strptime='%m/%Y')
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    sid = ''
    driver = False
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, ie, usuario, senha, perfil = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
        if usuario_anterior != usuario:
            # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
            if driver:
                driver.close()

            erro = 'S'
            contador = 0
            # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
            while erro == 'S':
                try:
                    # cookies, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil)
                    driver, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil, retorna_driver=True, orientacao='retrato')
                    erro = 'N'
                except:
                    print('❗ Erro ao logar na empresa, tentando novamente')
                    erro = 'S'
                    contador += 1
                
                sleep(1)
        
        driver, resultado = consulta_gia(ie, comp, driver, sid)
        
        # se a consulta retornar um ok, cria o PDF e captura os dados
        if resultado == 'ok':
            nome_arquivo = f'{cnpj} - Entrega de GIA {comp.replace("/", "-")}'
            driver, resultado = create_pdf(driver, nome_arquivo)
            driver, resultado = captura_dados(driver)
            
        # escreve na planilha de andamentos o resultado da execução atual
        try:
            _escreve_relatorio_csv(f"{cnpj};{ie};{resultado.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')}")
        except:
            raise Exception(f"Erro ao escrever esse texto: {resultado}")
        print(resultado)
        usuario_anterior = usuario
        
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;IE;NOME;DATA ENTREGA;RESPONSÁVEL;CHAVE DE AUTENTICAÇÃO;CNAE;ORIGEM;REFERÊNCIA;TIPO;CATEGORIA;PROTOCOLO')
    return True


if __name__ == '__main__':
    run()
