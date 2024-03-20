import pyautogui as p, PySimpleGUI as sg
import time
import os
import pyperclip
from threading import Thread

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def seleciona_certificado():
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('botao_gov.png', conf=0.9):
        time.sleep(1)
    
    # clica no botão do gov.br
    _click_img('botao_gov.png', conf=0.9)
    
    # aguarda o botão de certificado
    while not _find_img('certificado.png', conf=0.95):
        time.sleep(1)
        
    # clica no botão de certificado
    _click_img('certificado.png', conf=0.95)
    
    # aguarda a seleção de certificado
    timer = 0
    while not _find_img('seleciona_certificado.png'):
        if _find_img('captcha.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return 'recomeçar'
        time.sleep(1)
        timer += 1
        
        if timer > 10:
            _acessar_site_chrome('https://fgtsdigital.sistema.gov.br/portal/login')
            time.sleep(1)
            while not _find_img('botao_gov.png', conf=0.9):
                time.sleep(1)
            
            # clica no botão do gov.br
            _click_img('botao_gov.png', conf=0.9)
            
            # aguarda o botão de certificado
            while not _find_img('certificado.png', conf=0.95):
                if _find_img('erro_gov.png'):
                    p.press('pgDn')
                    time.sleep(1)
                    # clica no botão do gov.br
                    _click_img('botao_gov.png', conf=0.9)
                    
            # clica no botão de certificado
            _click_img('certificado.png', conf=0.95)
            timer = 0
            
    p.press('enter')


def login(cnpj):
    # aguarda a seleção de perfil
    while not _find_img('defina_perfil.png'):
        if _find_img('erro_gov.png'):
            p.press('pgDn')
            time.sleep(1)
            # clica no botão do gov.br
            _click_img('botao_gov.png', conf=0.9)
        time.sleep(1)
    
    if _find_img('aceita_cookies.png', conf=0.95):
        _click_img('aceita_cookies.png', conf=0.95)
    
    if not _find_img('procurador_selecionado.png', conf=0.95):
        # clica para definir perfil
        _click_img('seleciona_perfil.png', conf=0.9, timeout=1)
        
        # aguarda aparecer procurador
        while not _find_img('perfil_procurador.png'):
            if _find_img('procurador_cinza.png'):
                break
            time.sleep(1)
            
        # clica em procurador
        _click_img('perfil_procurador.png', conf=0.9, timeout=1)
    
    time.sleep(1)
    
    # clica no campo
    _click_position_img('informe_cnpj.png', '+', pixels_y=39, clicks=3, conf=0.95)
    time.sleep(1)
    
    p.write(cnpj)
    time.sleep(1)
    
    # clica em definir
    _click_img('definir.png', conf=0.9)
    time.sleep(1)
    
    if _find_img('sem_procuracao.png', conf=0.95):
        return 'Não existe procuração para o Número de Inscriçao selecionado'
    
    return 'ok'


def busca_guia():
    # aguarda aparecer o campo de CNPJ
    while not _find_img('gestao_guias.png', conf=0.9):
        if _find_img('cadastro_incompleto.png', conf=0.9):
            return 'Cadastro incompleto'
        time.sleep(1)
    
    _acessar_site_chrome('https://fgtsdigital.sistema.gov.br/cobranca/#/gestao-guias/emissao-guia-rapida')
    
    print('>>> Buscando guia')
    while not _find_img('emissao_de_guia.png', conf=0.9):
        time.sleep(1)
        
    if _find_img('sem_debitos.png', conf=0.95):
        _click_img('sem_debitos.png', conf=0.95, clicks=3)
        while True:
            try:
                p.hotkey('ctrl', 'c')
                p.hotkey('ctrl', 'c')
                mensagem = pyperclip.paste()
                break
            except:
                pass
        
        mensagem = mensagem.split('(')
        mensagem = f'{mensagem[0]};{mensagem[1].replace(")", "")}- {mensagem[2].replace(")", "")}'
        
        return mensagem.replace("\n", "").replace("\r", "")
    
    return 'Empresa ok'


@_time_execution
@_barra_de_status
def run(window):
    resultado = ''
    andamentos = f'Guias FGTS Digital'
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)
        
        cnpj, nome = empresa
        
        if resultado != 'Não existe procuração para o Número de Inscriçao selecionado':
            while True:
                _abrir_chrome('https://fgtsdigital.sistema.gov.br/portal/login')
                resultados = seleciona_certificado()
                if resultados != 'recomeçar':
                    break
                    
        resultado = login(cnpj)
        
        if resultado == 'Não existe procuração para o Número de Inscriçao selecionado':
            _escreve_relatorio_csv(f'{cnpj};{nome};{resultado}', nome=andamentos)
            continue
                
        resultado = busca_guia()
        
        _escreve_relatorio_csv(f'{cnpj};{nome};{resultado}', nome=andamentos)
        p.hotkey('ctrl', 'w')
        
        # Salvar a guia
        # imprimir(empresa, andamentos)
    
    time.sleep(2)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
