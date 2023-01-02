# -*- coding: utf-8 -*-
from PIL import Image
import os, time, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

os.makedirs('execução/documentos', exist_ok=True)


def login(cnpj, nome):
    time.sleep(2)
    p.hotkey('win', 'm')
    os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
    
    # espera o navegador abrir e clica na barra de endereços
    _wait_img('navegador.png', conf=0.9, timeout=-1)
    time.sleep(1)
    
    p.hotkey('alt', 'space', 'x')
    
    _click_img('navegador.png', conf=0.9)
    time.sleep(1)
    
    p.write('https://conectividadesocialv2.caixa.gov.br/sicns/')
    time.sleep(1)
    
    p.press('enter')
    
    # espera a página abrir e clica no botão empregador
    _wait_img('empregador.png', conf=0.9, timeout=-1)
    time.sleep(1)
    _click_img('empregador.png', conf=0.9)
    time.sleep(1)
    
    # espera a janela de certificados aparecer e confirma o RPEM
    _wait_img('seleciona_certificado.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.press('enter')
    
    # espera a tela inicial do sistema abrir se enquanto isso pedir o certificado de novo, confirma o RPEM
    while not _find_img('selecione_o_servico.png', conf=0.9):
        time.sleep(1)
        
        if _find_img('seleciona_certificado.png', conf=0.9):
            time.sleep(1)
            p.press('enter')

    # clica para selecionar o serviço desejado
    _click_img('selecione_o_servico.png', conf=0.9)


    # espera abrir o dropdown e clica em acessar empresa
    _wait_img('acessar_empresa_outorgante.png', conf=0.9, timeout=-1)
    time.sleep(1)
    _click_img('acessar_empresa_outorgante.png', conf=0.9)
    time.sleep(1)
    
    # espera abrir a tela de acesso empresa e clica no campo para inserir o CNPJ
    _wait_img('insere_cnpj.png', conf=0.9, timeout=-1)
    time.sleep(1)
    _click_img('insere_cnpj.png', conf=0.9)
    time.sleep(1)
    
    print('ok')
    time.sleep(55)
    

@_time_execution
def run():
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, pis = empresa
        
        _indice(count, total_empresas, empresa)
        
        if not login(cnpj, nome):
            break


if __name__ == '__main__':
    run()
