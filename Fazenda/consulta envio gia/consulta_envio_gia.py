import time, re, os
from pyautogui import prompt
from requests.exceptions import ConnectionError
from bs4 import BeautifulSoup
from sys import path
from requests import Session
from time import sleep
from PIL import Image
from fpdf import FPDF

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

path.append(r'..\..\_comum')
from pyautogui_comum import _get_comp
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from fazenda_comum import _get_info_post, _new_session_fazenda_driver
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _download_file, _open_lista_dados, _where_to_start, _indice


def captura_dados(driver):
    infos = re.compile(r'<td class=\"GIA-CABEC-CONSULTA-COMPLETA\">(.+)</td>').findall(driver.page_source)
    resultado = []
    for info in infos:
        resultado.append(info)
    
    tipo = re.compile(r'<td class=\"GIA-CABEC-CONSULTA-COMPLETA\">\n +	(.+)').search(driver.page_source).group(1)
    responsavel = re.compile(r'').search(driver.page_source).group(1)
    
    return driver, f'{resultado[1]};{resultado[2]};{resultado[3]};{tipo};{resultado[12]};{resultado[13]};{resultado[14]};{resultado[15]};{resultado[16]}'


def create_pdf(driver, nome_arquivo):
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
    print(f'>>> Consultando GIA')
    comp = comp.split('/')
    mes = comp[0]
    if mes[0] == '0': mes = mes[1]
    
    ano = comp[1]
    
    driver.get('https://cert01.fazenda.sp.gov.br/novaGiaEFDWEB/consultaCompleta.gia?method=init&SID=' + sid)
    
    while not _find_by_id('ie', driver):
        time.sleep(1)
    
    driver.find_element(by=By.ID, value='ie').send_keys(ie)
    
    drops = [('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/select[1]', mes),
             ('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/select[1]', mes),
             ('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/select[2]', ano),
             ('/html/body/form/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/select[2]', ano),
             ]
    
    for drop in drops:
        # Localize o elemento do dropdown
        dropdown_element = driver.find_element(by=By.XPATH, value=drop[0])
        # Crie um objeto Select com o elemento do dropdown
        dropdown = Select(dropdown_element)
        # Selecione o item desejado pelo valor
        dropdown.select_by_value(drop[1])
    
    time.sleep(5)
    driver.find_element(by=By.XPATH, value='/html/body/form/table/tbody/tr[4]/td/input').click()
    
    resultado = re.compile(r'RESULTADO-ERRO.+\n(.+)').search(driver.page_source)
    if resultado:
        return driver, resultado.group(1)
    
    driver.find_element(by=By.XPATH, value='/html/body/form/table[3]/tbody/tr[2]').click()
    
    return driver, 'ok'


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
                """if contador >= 3:
                    cookies = 'erro'
                    sid = 'Erro ao logar na empresa'
                    break"""
                try:
                    # cookies, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil)
                    driver, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil, retorna_driver=True)
                    erro = 'N'
                except:
                    print('❗ Erro ao logar na empresa, tentando novamente')
                    erro = 'S'
                    contador += 1
                
                sleep(1)
                """# se não salvar os cookies fecha a sessão e vai para o próximo dado
                if cookies == 'erro':
                    texto = f'{cnpj};{sid}'
                    usuario_anterior = 'padrão'
                    s.close()
                    _escreve_relatorio_csv(texto)
                    print(f'❗ {sid}\n', end='')
                    continue
                
                # adiciona os cookies do login da sessão por request no web driver
                for cookie in cookies:
                    s.cookies.set(cookie['name'], cookie['value'])"""
        
        """# se não retornar a id da sessão do web driver fecha a sessão por request
        if not sid:
            situacao = '❌ Erro no login'
            usuario_anterior = 'padrão'
            s.close()
        
        # se retornar a id da sessão do web driver executa a consulta
        else:
            # retorna o resultado da consulta
            situacao = consulta_gia(ie, comp, s, sid)
            # guarda o usuario da execução atual
            usuario_anterior = usuario"""
        
        driver, resultado = consulta_gia(ie, comp, driver, sid)
        
        if resultado == 'ok':
            nome_arquivo = f'{cnpj} - Entrega de GIA {comp.replace("/", "-")}'
            driver, resultado = create_pdf(driver, nome_arquivo)
            driver, resultado = captura_dados(driver)
            
        # escreve na planilha de andamentos o resultado da execução atual
        try:
            _escreve_relatorio_csv(f"{cnpj};{resultado.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')}")
        except:
            raise Exception(f"Erro ao escrever esse texto: {resultado}")
        print(resultado)
        usuario_anterior = usuario
        
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('CNPJ;CNAE;REGIME;SUBSTITUIÇÃO TRIBUTÁRIA;TIPO DA GIA;ORIGEM;PROTOCOLO;CONTROLE;CONTROLE CONTA FISCAL;DATA ENTREGA')
    return True


if __name__ == '__main__':
    run()
