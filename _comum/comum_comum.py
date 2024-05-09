# -*- coding: utf-8 -*-
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from requests import Session
from functools import wraps
from pathlib import Path
from pyautogui import alert, confirm
from threading import Thread
from io import BytesIO
from PIL import Image
import time, os, re, traceback, tempfile, contextlib, OpenSSL.crypto, PySimpleGUI as sg

# variáveis globais
e_dir = Path('execução')
_headers = {
    'Accept-Encoding': 'gzip, deflate, sdch',
    'Accept-Language': 'en-US,en;q=0.8',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
}


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
        window = sg.Window('', layout, no_titlebar=True, keep_on_top=True, element_justification='center', size=(745, 34), margins=(0,0), finalize=True, location=((screen_width // 2) - (745 // 2), 0))
        
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
                func(window)
            except Exception as e:
                # Obtém a pilha de chamadas de volta como uma string
                traceback_str = traceback.format_exc()
                
                window['-titulo-'].update(visible=False)
                window['-iniciar-'].update(visible=False)
                window['-Processando-'].update(visible=False)
                window['-Mensagens-'].update(visible=False)
                
                window['-fechar-'].update(visible=True)
                window['-titulo-'].update('Erro', text_color='#fc0303')
                window['-titulo-'].update(visible=True)
                window['-Error-'].update(visible=True)
                
                alert(f'Traceback: {traceback_str}\n\n'
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
_barra_de_status = barra_de_status


# Decorator que mede o tempo de execução de uma função decorada com ela
def time_execution(func):
    @wraps(func)
    def wrapper():
        comeco = datetime.now()
        print(f"🕐 Execução iniciada as: {comeco}\n")
        func()
        print(f"\n🕞 Tempo de execução: {datetime.now() - comeco}\n🕖 Encerrado as: {datetime.now()} ")
    return wrapper
_time_execution = time_execution


'''
Decorator para cenários onde pode ou não já existir uma instancia de BeatifulSoup
Cria uma instancia de BeautifulSoup, referente ao response.content 'content'
recebido, e passa para a função 'func' recebida ou simplesmente passa a
instancia de BeautifulSoup 'soup' recebida

Na declaração de uma funcao genérica utiliza-se a seguinte forma

@content_or_soup
def funcao_generica(soup):
    # faça alguma_coisa
    # return alguma_coisa ou não

Dessa forma qualquer funcao decorada com esse decorator pode receber tanto um
response.content ou uma instancia de BeautifulSoup da seguinte forma

para response.content -> funcao_generica(content=response.content)
para instancia BeautifulSoup -> funcao_generica(soup=BeautifulSoup(content, 'parser'))
'''


def content_or_soup(func):
    def wrapper(*dump, content=None, soup=None):
        if content:
            soup = BeautifulSoup(content, 'html.parser')
        return func(soup, *dump)

    return wrapper
_content_or_soup = content_or_soup


@contextlib.contextmanager
def pfx_to_pem(pfx_path, pfx_psw):
    # Decrypts the .pfx file to be used with requests.
    with tempfile.NamedTemporaryFile(suffix='.pem', delete=False) as t_pem:
        pfx = open(pfx_path, 'rb').read()  # --> type(pfx) > bytes
        p12 = OpenSSL.crypto.load_pkcs12(pfx, pfx_psw.encode('utf8'))

        if p12.get_certificate().has_expired(): 
            print('Certificado possivelmente vencido')
            dti = p12.get_certificate().get_notBefore()
            dtv = p12.get_certificate().get_notAfter()
            print(f'Inicio: {dti[6:8]}/{dti[4:6]}/{dti[:4]}')
            print(f'Vencimento: {dtv[6:8]}/{dtv[4:6]}/{dtv[:4]}')

        f_pem = open(t_pem.name, 'wb')
        f_pem.write(OpenSSL.crypto.dump_privatekey(OpenSSL.crypto.FILETYPE_PEM, p12.get_privatekey()))
        f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, p12.get_certificate()))
        ca = p12.get_ca_certificates()
        if ca is not None:
            for cert in ca:
                f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, cert))
        f_pem.close()

        yield t_pem.name
_pfx_to_pem = pfx_to_pem


# Recebe um cnpj e procura na pasta _cert por um certificado correspondente, caso 'senha is None' tenta os 8 primeiros dígitos do cnpj como senha
# Retorna o caminho do cert em caso de sucesso
# Retorna mensagem de erro caso falha
def which_cert(cnpj, psw=None):
    c_dir = os.path.join('..', '..', '_cert')
    
    pfxs = tuple(i for i in os.listdir(c_dir) if i[-4:] == '.pfx')

    for pfx in pfxs:
        if not psw:
            psw = re.search(r'(\d{8})', pfx)
            if psw:
                psw = psw.group(1)
            else:
                raise Exception('Falta senha no nome do cert', pfx)

        try:
            pfx = os.path.join(c_dir, pfx)
            with open(pfx, 'rb') as cert:
                p12 = OpenSSL.crypto.load_pkcs12(
                    cert.read(), psw.encode('utf8')
                )
        except OpenSSL.crypto.Error:
            continue

        cert_id = p12._friendlyname.decode()
        if cnpj in cert_id:
            return pfx, psw

    return 'cnpj e/ou senha não correspondentes'
_which_cert = which_cert


# Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'
def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()
_escreve_relatorio_csv = escreve_relatorio_csv


# Recebe um cabeçalho 'texto' e escreve
# no comeco do arquivo 'nome'
def escreve_header_csv(texto, nome='resumo', local=e_dir, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        with open(os.path.join(local, f"{nome}.csv"), 'r', encoding=encode) as f:
            conteudo = f.read()
        with open(os.path.join(local, f"{nome}.csv"), 'w', encoding=encode) as f:
            f.write(texto + '\n' + conteudo)
    except:
        with open(os.path.join(local, f"{nome} - auxiliar.csv"), 'r', encoding=encode) as f:
            conteudo = f.read()
        with open(os.path.join(local, f"{nome} - auxiliar.csv"), 'w', encoding=encode) as f:
            f.write(texto + '\n' + conteudo)
    
_escreve_header_csv = escreve_header_csv


# wrapper para askopenfilename
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


# wrapper para askdirectory
def ask_for_dir(title='Abrir pasta'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)

    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False
_ask_for_dir = ask_for_dir


# Procura pelo indice do primeiro campo da última linha, que deve ser um identificador unico como cpf/cnpj/inscricao, do arquivo de resumo escolhido e o retorna
# 
# Retorna 0 caso escolha não na caixa de texto ou não encontre o identificador na variável 'idents'
#
# Retorna None caso caixa de texto seja fechada ou o arquivo resumo, não seja escolhido ou não consiga abrir o arquivo escolhido
def where_to_start(idents, encode='latin-1'):
    title = 'Execucao anterior'
    text = 'Deseja continuar execucao anterior?'

    res = confirm(title=title, text=text, buttons=('sim', 'não'))
    if not res:
        return None
    if res == 'não':
        return 0

    ftypes = [('Plain text files', '*.txt *.csv')]
    file = ask_for_file(filetypes=ftypes)
    if not file:
        return None
    
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return None

    try:
        elem = dados[-1].split(';')[0]
        return idents.index(elem) + 1
    except ValueError:
        return 0
_where_to_start = where_to_start


# Abre uma janela de seleção de arquivos e abre o arquivo selecionado
# Retorna List de Tuple das linhas dividas por ';' do arquivo caso sucesso
# Retorna None caso nenhum selecionado ou erro ao ler arquivo
def open_lista_dados(i_dir='ignore', encode='latin-1', file=False):
    ftypes = [('Plain text files', '*.txt *.csv')]
    
    if not file:
        file = ask_for_file(filetypes=ftypes, initialdir=i_dir)
        if not file:
            return False

    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{i_dir}\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))
_open_lista_dados = open_lista_dados



# Recebe o 'response' de um request e salva o conteudo num arquivo 'name' na pasta 'pasta'
# Retorna True em caso de sucesso
# Levanta uma exceção em caso de erro
def download_file(name, response, pasta=str(e_dir / 'docs')):
    pasta = str(pasta).replace('\\', '/')
    os.makedirs(pasta, exist_ok=True)

    with open(os.path.join(pasta, name), 'wb') as arq:
        for i in response.iter_content(100000):
            arq.write(i)
_download_file = download_file


# Recebe o indice da empresa atual da consulta e a quantidade total de empresas
# Se não for a primeira empresa printa quantas faltam e pula 2 linhas
# Printa o indice da empresa atual mais as infos da mesma
def indice(count, total_empresas, empresa='', index=0, window=False, tempos=False, tempo_execucao=None):
    tempo_estimado_texto = ''
    tempo_estimado = 0
    if tempos:
        tempo_inicial = datetime.now()
        
        tempos.append(tempo_inicial)
        tempo_execucao_atual = int(tempos[1].timestamp()) - int(tempos[0].timestamp())
        
        tempo_execucao.append(tempo_execucao_atual)
        for t in tempo_execucao:
            tempo_estimado = tempo_estimado + t
        tempo_estimado = int(tempo_estimado) / int(len(tempo_execucao))
        
        tempo_total_segundos = int((len(total_empresas) + index) - (count + index) + 1) * int(tempo_estimado)
        tempo_estimado = tempo_execucao
        # Converter o tempo total para um objeto timedelta
        tempo_total = timedelta(seconds=tempo_total_segundos)
        
        # Extrair dias, horas e minutos do timedelta
        dias = tempo_total.days
        horas = tempo_total.seconds // 3600
        minutos = (tempo_total.seconds % 3600) // 60
        
        # Retorna o tempo no formato "dias:horas:minutos:segundos"
        tempo_estimado_texto = f" | Tempo estimado: {dias} dias {horas} horas {minutos} minutos"
        tempos.pop(0)
        
    if window:
        while True:
            try:
                window['-Mensagens-'].update(f'{str((count + index) - 1)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index) + 1)} Restantes{tempo_estimado_texto}')
                break
            except:
                try:
                    window['-Mensagens-'].update('Buguei...')
                except:
                    pass
                print('>>> Erro ao atualizar a interface, tentando novamente...')
                pass

    if count > 1:
        print(f'[ {len(total_empresas) - (count - 1)} Restantes ] {tempo_estimado_texto}\n\n')
    # Cria um indice para saber qual linha dos dados está
    indice_dados = f'[ {str(count + index)} de {str(len(total_empresas) + index)} ]'

    empresa = str(empresa).replace("('", '[ ').replace("')", ' ]').replace("',)", " ]").replace(',)', ' ]').replace("', '", ' - ')
            
    print(f'{indice_dados} - {empresa}')
    return tempos, tempo_estimado
_indice = indice


def remove_emojis(text):
    emoji_pattern = re.compile("["
                               u"\U0001F600-\U0001F64F"  # emoticons
                               u"\U0001F300-\U0001F5FF"  # símbolos e pictogramas
                               u"\U0001F680-\U0001F6FF"  # transportes e símbolos de mapa
                               u"\U0001F1E0-\U0001F1FF"  # bandeiras de país
                               "]+", flags=re.UNICODE)
    text = emoji_pattern.sub(r'', text)
    text = text[0:-1]
    return text
_remove_emojis = remove_emojis


def remove_espacos(text):
    string = re.compile(r"\s\s")
    text = string.sub(r'', text)
    text = text[0:-1]
    return text
_remove_espacos = remove_espacos


# gera um número aleatório de 10 dígitos que não contem na lista
def generate_random_number(lista_controle):
    while True:
        controle = str(random.randint(10**9, 10**10 - 1))
        if controle not in lista_controle:
            lista_controle.append(controle)
            return controle, lista_controle