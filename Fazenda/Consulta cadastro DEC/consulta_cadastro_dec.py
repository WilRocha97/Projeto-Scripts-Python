import re, random
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


def escreve_doc(texto, local='execução', nome='Pagina', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    for arq in os.listdir(local):
        try:
            os.remove(arq)
        except:
            pass
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()
    

def abre_dec():
    while True:
        _abrir_chrome('https://www.dec.fazenda.sp.gov.br/DEC/UCLogin/login.aspx', iniciar_com_icone=True)
        resultado = seleciona_certificado()
        if resultado == 'ok':
            break
        p.hotkey('ctrl', 'w')
    
    _acessar_site_chrome('https://www.dec.fazenda.sp.gov.br/DEC/UCCredenciamento/ConsultarCredenciamento.aspx')


def seleciona_certificado():
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('certificado_acesso.png', conf=0.9):
        time.sleep(1)
    
    time.sleep(random.randrange(1, 5))
    # clica no botão do gov.br
    _click_img('certificado_acesso.png', conf=0.9)
    
    # aguarda a seleção de certificado
    while not _find_img('seleciona_certificado.png'):
        time.sleep(1)
    
    time.sleep(random.randrange(1, 5))
    p.press('enter')
    
    # aguarda a seleção de perfil
    while not _find_img('tela_inicial.png', conf=0.9):
        if _find_img('erro_entrar.png', conf=0.9):
            return 'erro'
        time.sleep(1)
    
    return 'ok'


def captura_info(andamentos):
    # aguarda o cabeçalho do site
    while not _find_img('titulo_site.png', conf=0.9):
        if _find_img('erro_entrar.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            abre_dec()
        time.sleep(1)
        
    # aguarda aparecer a tabela com as empresas
    while not _find_img('situacao_cadastro.png', conf=0.9):
        time.sleep(1)
    
    print('>>> Capturando dados')
    p.hotkey('ctrl', 'a')
    
    while True:
        try:
            p.hotkey('ctrl', 'c')
            p.hotkey('ctrl', 'c')
            conteudo_pagina = pyperclip.paste()
            break
        except:
            pass
    
    lista_empresas = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)	(.+)	Enviar (.+)\n').findall(conteudo_pagina)
    for empresa in lista_empresas:
        _escreve_relatorio_csv(f'{empresa[0]};{empresa[1]};{empresa[2]}'.replace("\r", ""), nome=andamentos)


@_time_execution
@_barra_de_status
def run(window):
    
    window['-Mensagens-'].update('')
    
    andamentos = f'Consulta Situação DEC'
    
    abre_dec()
    
    paginas = 1
    while True:
        captura_info(andamentos)
        
        time.sleep(random.randrange(3, 10))
        proxima = '1'
        while not p.locateOnScreen(os.path.join('imgs', 'next.png'), confidence=0.95):
            if p.locateOnScreen(os.path.join('imgs', 'next_2.png'), confidence=0.95):
                proxima = '2'
                break
            if p.locateOnScreen(os.path.join('imgs', 'next_3.png'), confidence=0.95):
                proxima = '3'
                break
                
            _click_img('rodape.png', conf=0.95, timeout=1)
            p.press('pgDn')
            
        if p.locateOnScreen(os.path.join('imgs', 'ultima_pagina.png')):
            break
        
        if proxima == '1':
            p.click(p.locateCenterOnScreen(os.path.join('imgs', 'next.png'), confidence=0.95))
        elif proxima == '2':
            p.click(p.locateCenterOnScreen(os.path.join('imgs', 'next_2.png'), confidence=0.95))
        elif proxima == '3':
            p.click(p.locateCenterOnScreen(os.path.join('imgs', 'next_3.png'), confidence=0.95))
        
        paginas += 1
        p.screenshot().save('debug.png')
        
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
