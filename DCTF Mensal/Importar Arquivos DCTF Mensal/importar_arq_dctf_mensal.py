# -*- coding: utf-8 -*-
import traceback, re, pyperclip, shutil, time, os, pyautogui as p, PySimpleGUI as sg
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from threading import Thread
from pathlib import Path
from functools import wraps
from io import BytesIO
from PIL import Image

e_dir = Path('dados')


def barra_de_status(func):
    # Add your new theme colors and settings
    sg.LOOK_AND_FEEL_TABLE['tema'] = {'BACKGROUND': '#ffffff',
                                      'TEXT': '#000000',
                                      'INPUT': '#ffffff',
                                      'TEXT_INPUT': '#ffffff',
                                      'SCROLL': '#ffffff',
                                      'BUTTON': ('#000000', '#ffffff'),
                                      'PROGRESS': ('#ffffff', '#ffffff'),
                                      'BORDER': 0,
                                      'SLIDER_DEPTH': 0,
                                      'PROGRESS_DEPTH': 0}
    
    @wraps(func)
    def wrapper():
        def image_to_data(im):
            """
            Image object to bytes object.
            : Parameters
              im - Image object
            : Return
              bytes object.
            """
            with BytesIO() as output:
                im.save(output, format="PNG")
                data = output.getvalue()
            return data
        
        filename = 'V:/Setor Robô/Scripts Python/_comum/Assets/processando.gif'
        im = Image.open(filename)
        
        filename_check = 'V:/Setor Robô/Scripts Python/_comum/Assets/check.png'
        im_check = Image.open(filename_check)
        
        filename_error = 'V:/Setor Robô/Scripts Python/_comum/Assets/error.png'
        im_error = Image.open(filename_error)
        
        index = 0
        frames = im.n_frames
        im.seek(index)
        
        sg.theme('tema')  # Define o tema do PySimpleGUI
        # sg.theme_previewer()
        # Layout da janela
        layout = [
            [sg.Button('Iniciar', key='-iniciar-', border_width=0),
             sg.Button('Fechar', key='-fechar-', border_width=0, visible=False),
             sg.Text('', key='-titulo-'),
             sg.Text('', key='-Mensagens-'),
             sg.Image(data=image_to_data(im), key='-Processando-'),
             sg.Image(data=image_to_data(im_check), key='-Check-', visible=False),
             sg.Image(data=image_to_data(im_error), key='-Error-', visible=False)],
        ]
        
        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, keep_on_top=True, element_justification='center', size=(535, 34), margins=(0, 0), finalize=True, location=((screen_width // 2) - (535 // 2), 0))
        
        def run_script_thread():
            # habilita e desabilita os botões conforme necessário
            window['-titulo-'].update('Rotina automática, não interfira.', text_color='#fca103')
            window['-iniciar-'].update(visible=False)
            window['-fechar-'].update(visible=False)
            window['-Check-'].update(visible=False)
            window['-Error-'].update(visible=False)
            window['-Processando-'].update(visible=True)
            
            try:
                # Chama a função que executa o script
                func(window, event)
            except Exception as e:
                # Obtém a pilha de chamadas de volta como uma string
                traceback_str = traceback.format_exc()
                
                window['-titulo-'].update(visible=False)
                window['-iniciar-'].update(visible=False)
                window['-Processando-'].update(visible=False)
                
                window['-fechar-'].update(visible=True)
                window['-titulo-'].update('Erro', text_color='#fc0303')
                window['-titulo-'].update(visible=True)
                window['-Error-'].update(visible=True)
                
                p.alert(f'Traceback: {traceback_str}\n\n'
                      f'Erro: {e}')
                print(f'Traceback: {traceback_str}\n\n'
                      f'Erro: {e}')
                return
            
            # habilita e desabilita os botões conforme necessário
            window['-titulo-'].update(visible=False)
            window['-Processando-'].update(visible=False)
            window['-Mensagens-'].update('')
            
            window['-iniciar-'].update(visible=True)
            window['-fechar-'].update(visible=True)
            window['-titulo-'].update('Execução finalizada')
            window['-titulo-'].update(visible=True)
            window['-Check-'].update(visible=True)
        
        processando = 'não'
        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read(timeout=15)
            
            if event == sg.WIN_CLOSED:
                break
            elif event == '-fechar-':
                break
            
            elif event == '-iniciar-':
                # Cria uma nova thread para executar o script
                Thread(target=run_script_thread).start()
                processando = 'sim'
            
            if processando == 'sim':
                index = (index + 1) % frames
                im.seek(index)
                window['-Processando-'].update(data=image_to_data(im))
        
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
    try:
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
    except:
        return False


# Abre uma janela de seleção de arquivos e abre o arquivo selecionado
# Retorna List de Tuple das linhas dividas por ';' do arquivo caso sucesso
# Retorna None caso nenhum selecionado ou erro ao ler arquivo
def open_lista_dados(file=False, encode='latin-1'):
    if not file:
        ftypes = [('Plain text files', '*.txt *.csv')]
        
        file = ask_for_file(filetypes=ftypes)
        if not file:
            return False
    
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except:
        return False
    
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def ask_for_file(title='Abrir arquivo', filetypes='*', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    return file if file else False


def ask_for_dir(title='Abra a pasta com os arquivos ".RFB" para importar'):
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


def cria_dados(pasta_arquivos):
    planilha_de_dados = p.confirm(buttons=['Criar planilha com os arquivos na pasta selecionada', 'Selecionar planilha já existente'])
    if planilha_de_dados == 'Criar planilha com os arquivos na pasta selecionada':
        try:
            for dado in os.listdir(e_dir):
                os.remove(os.path.join(e_dir, dado))
        except:
            pass
        
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
    else:
        dados = open_lista_dados()

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
            return
        time.sleep(1)

    for imagen in ['unidade_nao_selecionada.png', 'unidade_nao_selecionada_2.png', 'unidade_nao_selecionada_3.png']:
        if find_img(imagen, conf=0.9):
            click_img(imagen, conf=0.9, clicks=2)

    if event == '-encerrar-':
        return

    p.press(unidade)
    time.sleep(0.5)

    if event == '-encerrar-':
        return

    p.press('tab')
    time.sleep(0.5)

    if event == '-encerrar-':
        return

    for count, subpasta in enumerate(pasta):
        if event == '-encerrar-':
            return
        if subpasta != pasta[0]:
            p.write(subpasta)
            time.sleep(2)
            click_position_img('pasta_selecionada_2.png', '+', pixels_x=15, conf=0.9, clicks=2)
            time.sleep(1)

            
def abrir_arquivo(event, arquivo, empresa, pasta_arquivos):
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
            return 'Arquivo inválido, ou não encontrado'
        if find_img('declaracao_ja_importada.png', conf=0.9):
            p.press('enter')
            auxiliar = ', Declaração já importada'
        if find_img('declaracao_nao_importada.png', conf=0.9):
            p.press('enter')
            wait_img('imprimir_relatorio.png', conf=0.9)
            p.hotkey('alt', 'i')
            if not imprimir_relatorio(event, empresa, pasta_arquivos):
                return ''
            p.hotkey('alt', 'f')
            move_arquivo_usado(pasta_arquivos, 'Arquivos Com Erros', arquivo)
            return 'Erro ao importar a Declaração, relatório de erros salvo'
            
    p.press('enter')
    move_arquivo_usado(pasta_arquivos, 'Arquivos Importados', arquivo)
    return f'Arquivo importado com sucesso{auxiliar}'


def imprimir_relatorio(event, empresa, pasta_arquivos):
    pasta_relatorio = os.path.join(pasta_arquivos, 'Relatórios de erro').replace('//', '/')
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(pasta_relatorio, exist_ok=True)

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
    
    pyperclip.copy(pasta_relatorio)
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


def move_arquivo_usado(pasta_arquivos, local, arquivo):
    pasta_arquivos_importados = os.path.join(pasta_arquivos, local)
    os.makedirs(pasta_arquivos_importados, exist_ok=True)

    shutil.move(os.path.join(pasta_arquivos, arquivo), os.path.join(pasta_arquivos_importados, arquivo))


@barra_de_status
def run(window, event):
    andamentos = 'Importar Arquivos DCTF Mensal'
    
    total_empresas = empresas[:]
    
    abrir_dctf_mensal(event, pasta_arquivos)
    
    window['-titulo-'].update('Rotina automática, não interfira.')
    for count, empresa in enumerate(empresas[:], start=1):
        # printa o indice da empresa que está sendo executada
        
        window['-Mensagens-'].update(f'{str(count - 1)} de {str(len(total_empresas))} | {str((len(total_empresas)) - count + 1)} Restantes')
        
        cnpj, arquivo = empresa
        resultado = abrir_arquivo(event, arquivo, empresa, pasta_arquivos)
        if resultado != '':
            escreve_relatorio_csv(f'{cnpj};{arquivo};{resultado}', nome=andamentos, local=pasta_arquivos)

        if count % 200 == 0:
            # a cada 200 arquivos, faz um backup
            cria_copia_de_seguranca()
            abrir_dctf_mensal(event, pasta_arquivos)
            
        if event == '-encerrar-':
            return

    cria_copia_de_seguranca()


if __name__ == '__main__':
    empresas = None
    while True:
        pasta_arquivos = ask_for_dir()
        if not pasta_arquivos:
            break

        empresas = cria_dados(pasta_arquivos)

        if empresas:
            break

        p.alert(text='Essa pasta não contem arquivos ".RFB"')

    if empresas:
        run()

    p.alert(text='Execução finalizada.')
