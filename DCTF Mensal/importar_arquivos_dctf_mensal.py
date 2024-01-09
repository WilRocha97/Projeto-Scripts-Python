# -*- coding: utf-8 -*-
import re, pyperclip, shutil, time, os, pyautogui as p, PySimpleGUI as sg
from tkinter.filedialog import askdirectory, Tk
from threading import Thread
from pathlib import Path
from functools import wraps
from sys import path

e_dir = Path('dados')


def barra_de_status(func):
    @wraps(func)
    def wrapper():
        sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
        # sg.theme_previewer()
        # Layout da janela
        layout = [
            [sg.Text('Rotina automática em execução, por favor não interfira.', key='-titulo-', text_color='#fca400'),
             sg.Button('Iniciar', key='-iniciar-', border_width=0),
             sg.Button('Encerrar', key='-encerrar-', border_width=0),
             sg.Text('', key='-Mensagens-')],
        ]
        
        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, keep_on_top=True, element_justification='center', size=(650, 35), border_depth=0, finalize=True, location=((screen_width // 2) - (600 // 2), 0))
        
        # window.move(screen_width // 2 - window.size[0] // 2, 0)
        
        def run_script_thread():
            # habilita e desabilita os botões conforme necessário
            window['-iniciar-'].update(disabled=True)
            
            try:
                # Chama a função que executa o script
                func(window, event)
            except Exception as erro:
                print(erro)
                
            # habilita e desabilita os botões conforme necessário
            window['-iniciar-'].update(disabled=False)
            
            # apaga qualquer mensagem na interface
            window['-Mensagens-'].update('')
            
        
        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read()
            
            if event == sg.WIN_CLOSED:
                break
            
            elif event == '-encerrar-':
                break
                
            elif event == '-iniciar-':
                # Cria uma nova thread para executar o script
                script_thread = Thread(target=run_script_thread)
                script_thread.start()
        
        window.close()
    
    return wrapper


def find_img(img, pasta='imgs', conf=1.0):
    try:
        path = os.path.join(pasta, img)
        return p.locateOnScreen(path, confidence=conf)
    except:
        return False


# Espera pela imagem 'img' que atenda ao nível de correspondência 'conf'
# Até que o 'timeout' seja excedido ou indefinidamente para 'timeout' negativo
# Retorna uma tupla com os valores (x, y, altura, largura) caso ache a img
# Retorna None caso não ache a imagem ou exceda o 'timeout'
def wait_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, debug=False):
    if debug:
        print('\tEsperando', img)

    aux = 0
    while True:
        box = find_img(img, pasta, conf=conf)
        if box:
            return box
        time.sleep(delay)

        if timeout < 0:
            continue
        if timeout == aux:
            break
        aux += 1

    return None


def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    try:
        for i in range(timeout):
            box = p.locateCenterOnScreen(img, confidence=conf)
            if box:
                p.click(p.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
                return True
            time.sleep(delay)
        else:
            return False
    except:
        return False
    
    
def click_position_img(img, operacao, pixels_x=0, pixels_y=0, pasta='imgs', conf=1.0, clicks=1):
    img = os.path.join(pasta, img)
    p.moveTo(p.locateCenterOnScreen(img, confidence=conf))
    local_mouse = p.position()
    if operacao == '+':
        p.click(int(local_mouse[0] + int(pixels_x)), int(local_mouse[1] + int(pixels_y)), clicks=clicks)
        return True
    if operacao == '-':
        p.click(int(local_mouse[0] - int(pixels_x)), int(local_mouse[1] - int(pixels_y)), clicks=clicks)
        return True
    if operacao == '+x-y':
        p.click(int(local_mouse[0] + int(pixels_x)), int(local_mouse[1] - int(pixels_y)), clicks=clicks)
        return True
    if operacao == '-x+y':
        p.click(int(local_mouse[0] - int(pixels_x)), int(local_mouse[1] + int(pixels_y)), clicks=clicks)
        return True


# Abre uma janela de seleção de arquivos e abre o arquivo selecionado
# Retorna List de Tuple das linhas dividas por ';' do arquivo caso sucesso
# Retorna None caso nenhum selecionado ou erro ao ler arquivo
def open_lista_dados(file, encode='latin-1'):
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        return False
    
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def ask_for_dir(title='Abrir pasta com os arquivos para importar'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False


# Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'
def escreve_relatorio_csv(texto, nome='dados', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


def cria_dados(window, pasta_arquivos):
    window['-Mensagens-'].update('Criando planilha de dados')

    for dado in os.listdir(e_dir):
        os.remove(os.path.join(e_dir, dado))
    
    arq_name = []
    for arq in os.listdir(pasta_arquivos):
        if arq.endswith('.RFB'):
            with open(os.path.join(pasta_arquivos, arq), 'r') as arquivo:
                # Leia o conteúdo do arquivo
                conteudo = arquivo.read()
                cnpj = re.compile(r'DCTFM\s+\d\d\d\d\d\d\d\d\d(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)\d\d\d\d').search(conteudo).group(1)
            
            names = arq.split('_')
            codigo = str(names[1]).replace('.RFB', '')
            arq_name.append((cnpj, codigo))
        
    arq_name = sorted(arq_name)
    for name in arq_name:
        escreve_relatorio_csv(f'{name[0]};DCTF_{name[1]}.RFB')
        
    dados = open_lista_dados(os.path.join(e_dir, 'dados.csv'))

    window['-Mensagens-'].update('')
    return dados
    
    
def abrir_dctf_mensal(event, pasta_arquivos):
    pasta = pasta_arquivos.split('/')
    unidade = pasta[0].replace(':', '')
    
    while not find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)

    click_img('tela_inicial.png', conf=0.9)

    time.sleep(1)
    while not find_img('origem_dos_dados.png', conf=0.9):
        p.hotkey('ctrl', 'm')
        if event == '-encerrar-':
            return ''
        time.sleep(1)

    for imagen in ['unidade_nao_selecionada.png', 'unidade_nao_selecionada_2.png', 'unidade_nao_selecionada_3.png']:
        if find_img(imagen, conf=0.9):
            click_img(imagen, conf=0.9, clicks=2)

    if event == '-encerrar-':
        return ''

    p.press(unidade)
    time.sleep(0.5)

    if event == '-encerrar-':
        return ''

    p.press('tab')
    time.sleep(0.5)

    if event == '-encerrar-':
        return ''

    for count, subpasta in enumerate(pasta):
        if event == '-encerrar-':
            return ''
        if subpasta != pasta[0]:
            p.write(subpasta)
            time.sleep(2)
            click_position_img('pasta_selecionada_2.png', '+', pixels_x=15, conf=0.9, clicks=2)
            time.sleep(1)

            
def abrir_arquivo(event, empresa, pasta_arquivos):
    while not find_img('nome_arquivo.png', conf=0.9):
        if event == '-encerrar-':
            return ''
        time.sleep(1)
    
    # clica na barra para digitar o nome do arquivo
    click_position_img('nome_arquivo.png', '+', pixels_y=17, conf=0.9, clicks=2)
    time.sleep(0.5)
    
    # escreve o nome do arquivo
    p.write(empresa[1])
    time.sleep(0.5)
    
    # aperta ok
    p.hotkey('alt', 'o')
    
    auxiliar = ''
    while not find_img('declaracao_importada.png', conf=0.9):
        if find_img('arquivo_invalido.png', conf=0.9):
            p.press('enter')
            return 'Arquivo inválido'
        if find_img('declaracao_ja_importada.png', conf=0.9):
            p.press('enter')
            auxiliar = 'Declaração já importada'
        if find_img('declaracao_nao_importada.png', conf=0.9):
            p.press('enter')
            wait_img('imprimir_relatorio.png', conf=0.9)
            p.hotkey('alt', 'i')
            if not imprimir_relatorio(event, empresa, pasta_arquivos):
                return ''
            p.hotkey('alt', 'f')
            return 'Erro ao importar a Declaração, relatório salvo'
            
    p.press('enter')
    return f'Arquivo importado com sucesso, {auxiliar}'


def imprimir_relatorio(event, empresa, pasta_arquivos):
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(os.path.join(pasta_arquivos, 'Relatórios de erro'), exist_ok=True)

    cnpj, arquivo = empresa
    arquivo = arquivo.replace('.RFB', '')
    
    while not find_img('salvar_como.png', conf=0.9):
        if event == '-encerrar-':
            return False
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    p.write(f'{cnpj} - {arquivo} - Relatórios de erro ')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.press('enter')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    pyperclip.copy(pasta_arquivos + '/Relatórios de erro')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.press('enter')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.hotkey('alt', 'l')
    time.sleep(1)
    
    if event == '-encerrar-':
        return False
    
    if find_img('substituir.png', conf=0.9):
        p.press('s')
    time.sleep(1)
    
    return True
    

def cria_copia_de_seguranca():
    os.makedirs('C:\Cópias DCTF', exist_ok=True)

    # fecha a janela de importação
    p.hotkey('alt', 'c')
    time.sleep(1)
    
    # abre o dropdown de ferramentas
    p.hotkey('alt', 'f')
    time.sleep(1)
    
    # abre a tela para salvar cópia de segurança
    p.press('s')
    time.sleep(1)

    # espera a tela abrir
    while not find_img('tela_salvar_copia.png', conf=0.9):
        time.sleep(1)

    # clica para selecionar a pasta de backup
    click_position_img('unidade_c_backup.png', '+', pixels_y=17, conf=0.99, clicks=2)
    time.sleep(1)

    # seleciona a unidade c
    p.press('c')

    # clica para selecionar a pasta de backup
    click_position_img('pasta_c_backup.png', '+', pixels_y=17, conf=0.99)
    time.sleep(1)

    # digita o nome da pasta
    p.write('Cópias DCTF')
    time.sleep(1)

    click_img('copias_dctf.png', conf=0.99, clicks=2)
    time.sleep(1)

    # seleciona todos os arquivos
    p.hotkey('alt', 't')
    time.sleep(1)

    # confirma
    p.hotkey('alt', 'o')
    time.sleep(1)

    # espera a tela abrir
    while not find_img('backup_criado.png', conf=0.9):
        if find_img('substituir_backup.png', conf=0.9):
            p.press('enter')

    time.sleep(1)

    p.press('enter')
    time.sleep(2)
    # confirma
    p.hotkey('alt', 'c')
    time.sleep(2)


def move_arquivo_usado(pasta_arquivos, arquivo):
    pasta_arquivos_importados = os.path.join(pasta_arquivos, 'Arquivos Importados')
    os.makedirs(pasta_arquivos_importados, exist_ok=True)

    shutil.move(os.path.join(pasta_arquivos, arquivo), os.path.join(pasta_arquivos_importados, arquivo))


@barra_de_status
def run(window, event):
    pasta_arquivos = ask_for_dir()
    empresas = cria_dados(window, pasta_arquivos)

    if not empresas:
        p.alert(text='Essa pasta não contem arquivos ".RFB"')

    andamentos = 'Importar Arquivos DCTF Mensal'

    total_empresas = empresas[:]
    
    abrir_dctf_mensal(event, pasta_arquivos)
    
    for count, empresa in enumerate(empresas[:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count)} de {str(len(total_empresas))} | {str((len(total_empresas)) - count)} Restantes')
        
        cnpj, arquivo = empresa
        resultado = abrir_arquivo(event, empresa, pasta_arquivos)
        if resultado != '':
            escreve_relatorio_csv(f'{cnpj};{arquivo};{resultado}', nome=andamentos, local=pasta_arquivos)

        move_arquivo_usado(pasta_arquivos, arquivo)

        if count % 200 == 0:
            # a cada 200 arquivos, faz um backup
            cria_copia_de_seguranca()
            abrir_dctf_mensal(event, pasta_arquivos)
            
        if event == '-encerrar-':
            return

    cria_copia_de_seguranca()


if __name__ == '__main__':
    run()
