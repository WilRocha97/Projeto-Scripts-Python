# -*- coding: utf-8 -*-
import pyperclip, os, time, pandas as pd, pyautogui as p, PySimpleGUI as sg
from pathlib import Path
from functools import wraps
from threading import Thread

mensagem_arquivo = "Dados\Mensagem.txt"
f = open(mensagem_arquivo, 'r', encoding='utf-8')
mensagem = f.read()


def barra_de_status(func):
    @wraps(func)
    def wrapper():
        sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
        # sg.theme_previewer()
        # Layout da janela
        layout = [
            [sg.Text('Rotina automática em execução, por favor não interfira.', key='-titulo-', text_color='#fca400'),
             sg.Button('Gerencia dados', key='-gerenciar-', border_width=0),
             sg.Button('Iniciar', key='-iniciar-', border_width=0),
             sg.Button('Fechar', key='-fechar-', border_width=0),
             sg.Text('', key='-Mensagens-')],
        ]
        
        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, element_justification='center', no_titlebar=True, location=(0, 0), size=(screen_width, 35), keep_on_top=True)
        
        def run_script_thread():
            try:
                # habilita e desabilita os botões conforme necessário
                window['-iniciar-'].update(disabled=True)
                
                # Chama a função que executa o script
                func(window, event)
                
                # habilita e desabilita os botões conforme necessário
                window['-iniciar-'].update(disabled=False)
                
                # apaga qualquer mensagem na interface
                window['-Mensagens-'].update('')
            except:
                pass
        
        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read()
            
            if event == sg.WIN_CLOSED:
                break
            
            elif event == '-gerenciar-':
                os.startfile(r"C:\Program Files (x86)\Automações\Envia anexo WP\Dados")
                
            elif event == '-fechar-':
                break
            
            elif event == '-iniciar-':
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()
        
        window.close()
    
    return wrapper


def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    for i in range(timeout):
        box = p.locateCenterOnScreen(img, confidence=conf)
        if box:
            p.click(p.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
            return True
        sleep(delay)
    else:
        return False


def find_img(img, pasta='imgs', conf=1.0):
    try:
        path = os.path.join(pasta, img)
        return p.locateOnScreen(path, confidence=conf)
    except:
        return False


def abrir_chrome(url):
    def abrir_nova_janela():
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        
        timer = 0
        while not find_img('google.png', conf=0.9):
            time.sleep(1)
            timer += 1
            if timer > 30:
                return 'sair'
            if find_img('restaurar_pagina.png', conf=0.9):
                click_img('restaurar_pagina.png', conf=0.9)
                time.sleep(1)
                p.press('esc')
                time.sleep(1)
    
    abrir_nova_janela()
    
    click_img('google.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 'space', 'x')
    time.sleep(1)
    
    p.click(1000, 51)
    time.sleep(0.5)
    p.hotkey('ctrl', 'a')
    time.sleep(0.5)
    p.press('delete')
    time.sleep(0.2)
    pyperclip.copy(url)
    time.sleep(0.2)
    p.hotkey('ctrl', 'v')
    
    time.sleep(1)
    p.press('enter')
    return 'ok'


def enviar_anexo(numero, imagem_anexo):
    timer = 0
    while not find_img('nova_conversa.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'
        
    click_img('nova_conversa.png', conf=0.9)
    time.sleep(1)
    
    p.write(str(numero))
    time.sleep(2)
    
    if find_img('sem_contato.png', conf=0.9):
        time.sleep(1)
        p.hotkey('ctrl', 'w')
        time.sleep(1)
        return 'Número não encontrado'
    
    time.sleep(1)
    
    timer = 0
    while find_img('procurando_numero.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'

    p.press('enter')
    timer = 0
    while not find_img('anexar.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'
    
    click_img('anexar.png', conf=0.9)
    time.sleep(1)
    timer = 0
    while not find_img('imagem.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'
    
    click_img('imagem.png', conf=0.9)
    timer = 0
    while not find_img('abrir.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'
    
    time.sleep(1)
        
    pyperclip.copy(imagem_anexo)
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')

    time.sleep(1)
    timer = 0
    while not find_img('digitar_mensagem.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 30:
            return 'sair'
        
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    pyperclip.copy(mensagem)
    
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    
    p.press('enter')
    
    time.sleep(5)

    p.press('f5')
    time.sleep(1)
    
    return 'ok'


@barra_de_status
def run(window, event):
    imagem_anexo = 'C:\Program Files (x86)\Automações\Envia anexo WP\Dados\Anexo.jpeg'
    try:
        df_andamentos = pd.read_excel(os.path.join('Dados', 'Andamentos.xlsx'), index_col=False)
    except:
        df_andamentos = pd.DataFrame(columns=['Número', 'Nome', 'Resultado'])
        df_andamentos.to_excel(os.path.join('Dados', 'Andamentos.xlsx'), index=False)
        df_andamentos = pd.read_excel(os.path.join('Dados', 'Andamentos.xlsx'), index_col=False)
        
    df = pd.read_excel(os.path.join('Dados', 'Contatos.xlsx'), index_col=False)
    total = list(df.itertuples())
    if len(total) == 0:
        p.alert(text='Planilha de contatos vazia.')
        return
    
    window['-Mensagens-'].update(f'|Total: {str(len(total))}')
    
    resultado = abrir_chrome('https://web.whatsapp.com/')
    if resultado == 'sair':
        return
    
    for count, key in enumerate(df.itertuples()):
        numero, nome = key[1], key[2]
        
        resultado = enviar_anexo(numero, imagem_anexo)
        
        if resultado == 'sair':
            return
        
        df_andamentos = df_andamentos._append({'Número': numero, 'Nome': nome, 'Resultado': resultado}, ignore_index=True)
        df_andamentos.to_excel(os.path.join('Dados', 'Andamentos.xlsx'), index=False)
        
        df.drop(key[0], inplace=True)
        df.to_excel(os.path.join('Dados', 'Contatos.xlsx'), index=False)
        
        window['-Mensagens-'].update(f'|Total: {str(len(total))}   |Faltam: {str(len(list(df.itertuples())))}   |Feitos: {str(count + 1)}')
        if event == '-fechar-':
            return
        
 
if __name__ == '__main__':
    run()