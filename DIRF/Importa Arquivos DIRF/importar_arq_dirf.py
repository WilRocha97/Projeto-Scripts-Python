# -*- coding: utf-8 -*-
import re, pyperclip, shutil, time, os, pyautogui as p, PySimpleGUI as sg
from datetime import datetime
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
                                      'PROGRESS_DEPTH': 0, }
    
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
        
        filename = 'Assets/processando.gif'
        im = Image.open(filename)
        
        filename_check = 'Assets/check.png'
        im_check = Image.open(filename_check)
        
        index = 0
        frames = im.n_frames
        im.seek(index)
        
        sg.theme('tema')  # Define o tema do PySimpleGUI
        # sg.theme_previewer()
        # Layout da janela
        layout = [
            [sg.Button('Iniciar', key='-iniciar-', border_width=0),
             sg.Button('Encerrar', key='-encerrar-', border_width=0),
             sg.Text('', key='-titulo-', text_color='#fca103'),
             sg.Text('', key='-Mensagens-'),
             sg.Image(data=image_to_data(im), key='-Processando-'),
             sg.Image(data=image_to_data(im_check), key='-Check-', visible=False)],
        ]
        
        # guarda a janela na variável para manipula-la
        screen_width, screen_height = sg.Window.get_screen_size()
        window = sg.Window('', layout, no_titlebar=True, keep_on_top=True, size=(600, 35), margins=(0,0), element_justification='center', finalize=True, location=((screen_width // 2) - (600 // 2), 0))
    
        def run_script_thread():
            # habilita e desabilita os botões conforme necessário
            window['-iniciar-'].update(disabled=True)
            window['-Processando-'].update(visible=True)
            window['-Check-'].update(visible=False)
            
            try:
                # Chama a função que executa o script
                func(window, event)
            except Exception as erro:
                print(erro)
                
            # habilita e desabilita os botões conforme necessário
            window['-iniciar-'].update(disabled=False)
            # apaga qualquer mensagem na interface
            window['-Mensagens-'].update('')
            window['-Processando-'].update(visible=False)
            window['-Check-'].update(visible=True)
            
        processando = 'não'
        while True:
            # captura o evento e os valores armazenados na interface
            event, values = window.read(timeout=15)
            
            if event == sg.WIN_CLOSED:
                break
            
            elif event == '-encerrar-':
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
# Retorna uma tupla com os valores (x, y, altura, largura) caso ache a img
def wait_img(img, pasta='imgs', conf=1.0, delay=1, debug=False):
    if debug:
        print('\tEsperando', img)

    aux = 0
    while True:
        box = find_img(img, pasta, conf=conf)
        if box:
            return box
        time.sleep(delay)

        aux += 1

    return None


def click_img(img, pasta='imgs', conf=1.0, delay=1, button='left', clicks=1):
    img = os.path.join(pasta, img)
    try:
        box = p.locateCenterOnScreen(img, confidence=conf)
        if box:
            p.click(p.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
            return True
        time.sleep(delay)
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


def ask_for_dir(title='Abra a pasta com os arquivos ".txt" para importar'):
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


def cria_dados(window, planilha_de_dados, pasta_arquivos):
    if planilha_de_dados == 'Criar planilha com os arquivos na pasta selecionada':
        window['-Mensagens-'].update('Criando planilha de dados...')

        try:
            for dado in os.listdir(e_dir):
                os.remove(os.path.join(e_dir, dado))
        except:
            pass
        
        arq_name = []
        for arq in os.listdir(pasta_arquivos):
            if arq.endswith('.txt'):
                with open(os.path.join(pasta_arquivos, arq), 'r', encoding="latin-1") as arquivo:
                    # Leia o conteúdo do arquivo
                    conteudo = arquivo.read()
                    for tipo in [r'DECPJ\|(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)\|(.+)\|\d\|', r'DECPF\|(\d\d\d\d\d\d\d\d\d\d\d)\|(.+)(\|N){5}']:
                        try:
                            info = re.compile(tipo).search(conteudo)
                            cnpj = info.group(1)
                            nome = info.group(2)
                            break
                        except:
                            pass
                        
                arq_name.append((cnpj, arq, nome))
            
        arq_name = sorted(arq_name)
        for name in arq_name:
            escreve_relatorio_csv(f'{name[0]};{name[1]};{name[2]}')
            
        dados = open_lista_dados(os.path.join(e_dir, 'dados.csv'))
    else:
        dados = open_lista_dados()

    return dados
    
    
def abrir_dirf(event):
    while not find_img('tela_inicial.png', conf=0.9):
        if event == '-encerrar-':
            return
        time.sleep(1)

    click_img('tela_inicial.png', conf=0.9)
    time.sleep(1)


def busca_arquivo(event, arquivo, pasta_arquivos):
    timer = 0
    while not find_img('nome_arquivo.png', conf=0.9):
        p.hotkey('alt', 'i')
        if event == '-encerrar-':
            return ''
        time.sleep(1)
        
        timer += 1
        if timer > 10:
            p.click(p.locateCenterOnScreen(r'imgs\buscar_arquivo.png', confidence=0.95))

    # escreve o nome do arquivo
    while True:
        try:
            pyperclip.copy(str(os.path.join(pasta_arquivos, arquivo)))
            pyperclip.copy(str(os.path.join(pasta_arquivos, arquivo)))
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    
    time.sleep(0.5)
    
    # aperta ok
    p.press('enter')
    time.sleep(2)

    
def abrir_arquivo(event, cnpj, arquivo, empresa, pasta_arquivos, pasta_arquivos_importados, pasta_arquivos_nao_importados, pasta_arquivos_importados_com_erros_avisos, pasta_relatorios, pasta_log):
    busca_arquivo(event, arquivo, pasta_arquivos)
    
    timer = 0
    auxiliar = ''
    print('>>> Aguardando importar')
    while not p.locateOnScreen(r'imgs\arquivo_selecionado.png'):
        if find_img('erros_avisos.png', conf=0.9):
            print('>>> Foram encontrados erros e/ou avisos')
            p.hotkey('alt', 'n')
            resultado, mensagem = imprimir_relatorio(event, empresa, pasta_relatorios)
            if resultado == 'erro critico':
                return resultado
            if resultado == 'erro':
                move_arquivo_usado(pasta_arquivos, pasta_arquivos_nao_importados, arquivo, verifica=pasta_arquivos_importados)
                move_arquivo_log(pasta_log, arquivo)
                return f'Arquivo não importado{auxiliar}{mensagem}, arquivos de log salvos'
            
            auxiliar += ';Declaração com erros ou avisos, relatório salvo'
            move_arquivo_usado(pasta_arquivos, pasta_arquivos_nao_importados, arquivo, verifica=pasta_arquivos_importados)
            p.hotkey('alt', 'f4')
            return f'Arquivo não importado{auxiliar}'
        
        if find_img('tela_selecao.png', conf=0.9):
            return ''
        if p.locateOnScreen(r'imgs\arquivo_selecionado_2.png'):
            break
        time.sleep(1)
        
        if timer > 10:
            busca_arquivo(event, arquivo, pasta_arquivos)
        
        if find_img('declaracao_ja_importada.png', conf=0.9):
            while find_img('declaracao_ja_importada.png', conf=0.9):
                click_img('declaracao_ja_importada.png', conf=0.9)
                p.press('enter')
                time.sleep(1)
            auxiliar = ', Declaração já importada'
            break
            
        if event == '-encerrar-':
            return ''
        
        if find_img('arquivo_invalido.png', conf=0.9):
            p.press('enter')
            time.sleep(0.5)
            p.press('esc')
            return 'Arquivo inválido, ou não encontrado'
        timer += 1
        
    print('>>> Confirmando importação')
    p.hotkey('alt', 'a')
    time.sleep(0.5)
    
    while not find_img('resumo_importacao.png', conf=0.9):
        if find_img('tela_selecao.png', conf=0.9):
            return ''
        
        if find_img('declaracao_ja_importada.png', conf=0.9):
            while find_img('declaracao_ja_importada.png', conf=0.9):
                click_img('declaracao_ja_importada.png', conf=0.9)
                p.press('enter')
                time.sleep(1)
            auxiliar = ', Declaração já importada'
            p.hotkey('alt', 'a')
        
        if find_img('ja_consta_info_alimentandos.png', conf=0.9):
            p.hotkey('alt', 's')
            
        if find_img('aplicar_novamente.png', conf=0.9):
            p.hotkey('alt', 'p')
            time.sleep(0.5)
            p.hotkey('alt', 'i')
            time.sleep(0.5)
            if find_img('nome_arquivo.png', conf=0.9):
                p.hotkey('alt', 'f4')
            if p.locateCenterOnScreen(r'imgs\buscar_arquivo.png', confidence=0.95):
                p.hotkey('alt', 'f4')
            
            p.hotkey('alt', 'a')
    
    if find_img('erros_avisos.png', conf=0.9):
        print('>>> Foram encontrados erros e/ou avisos')
        p.hotkey('alt', 'n')
        resultado, mensagem = imprimir_relatorio(event, empresa, pasta_relatorios)
        
        if resultado == 'erro critico':
            return resultado
        if resultado == 'erro':
            move_arquivo_usado(pasta_arquivos, pasta_arquivos_importados_com_erros_avisos, arquivo, verifica=pasta_arquivos_importados)
            move_arquivo_log(pasta_log, arquivo)
            return f'Arquivo importado com sucesso{auxiliar}{mensagem}, arquivos de log salvos'
        
        p.hotkey('alt', 'f4')
        """time.sleep(3)
        if find_img('nao_importou.png', conf=0.9):
            auxiliar += ';Declaração com erros ou avisos, relatório salvo'
            move_arquivo_usado(pasta_arquivos, pasta_arquivos_nao_importados, arquivo, verifica=pasta_arquivos_importados)
            return f'Arquivo não importado{auxiliar}'"""
        
        auxiliar += ';Declaração com erros ou avisos, relatório salvo'
        move_arquivo_usado(pasta_arquivos, pasta_arquivos_importados_com_erros_avisos, arquivo, verifica=pasta_arquivos_importados)
        return f'Arquivo importado com sucesso{auxiliar}'

    p.hotkey('alt', 'n')
    move_arquivo_usado(pasta_arquivos, pasta_arquivos_importados, arquivo, verifica=pasta_arquivos_importados_com_erros_avisos)
    deletar_relatorios_antigos(cnpj, arquivo, pasta_relatorios)
    auxiliar += ';Declaração ok'
    return f'Arquivo importado com sucesso{auxiliar}'


def imprimir_relatorio(event, empresa, pasta_relatorios):
    print('>>> Aguardando tela para salvar relatório')
    timer = 0
    while not find_img('imprimir.png', conf=0.9):
        if find_img('caracteres_ocultos.png', conf=0.9):
            click_img('caracteres_ocultos.png', conf=0.9)
            p.press('enter')
            return 'erro', ';Não é possível gerar o relatório de importação, pois o arquivo contém caracteres ocultos não permitidos'
        time.sleep(1)
        timer += 1
        if timer > 20:
            p.hotkey('alt', 'f4')
            return 'erro critico', ';Erro ao imprimir o relatório de erros'
    
    print('>>> Salvando relatório')
    while True:
        while not find_img('salvar_em.png', conf=0.9):
            click_img('imprimir.png', conf=0.99)
            p.moveTo(10, 10)
            time.sleep(1)
        
        click_img('desktop.png', conf=0.99)
        
        cnpj, arquivo, nome = empresa
        arquivo = arquivo.replace('.txt', '')
        
        time.sleep(0.5)
        nome_arquivo = f'{cnpj} - {arquivo} - Relatório de erro.pdf'
        
        while find_img('salvar_em.png', conf=0.9):
            if find_img('digita_nome_arquivo.png', conf=0.9):
                click_position_img('digita_nome_arquivo.png', '+', pixels_x=100, conf=0.9, clicks=2)
            if find_img('digita_nome_arquivo_2.png', conf=0.9):
                click_position_img('digita_nome_arquivo_2.png', '+', pixels_x=230, conf=0.9, clicks=2)
            p.press('del', presses=50)
            p.press('backspace', presses=50)
            time.sleep(0.5)
            # digita o caminho para salvar o arquivo
            while True:
                try:
                    pyperclip.copy(nome_arquivo)
                    pyperclip.copy(nome_arquivo)
                    p.hotkey('ctrl', 'v')
                    break
                except:
                    pass
        
            time.sleep(0.5)
            
            if event == '-encerrar-':
                return False, ''
            
            p.press('enter')
            
        if find_img('substituir.png', conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f4')
        time.sleep(1)
        
        if find_img('erro_imprimir.png', conf=0.9):
            click_img('erro_imprimir.png', conf=0.9)
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'f4')
            return 'erro critico', ';Erro ao imprimir o relatório de erros'
        
        os.makedirs(pasta_relatorios, exist_ok=True)
        try:
            shutil.move(os.path.join(r'C:\Users\robo\Desktop', nome_arquivo), os.path.join(pasta_relatorios, nome_arquivo))
            break
        except:
            pass
        
    return True, ''
    

def cria_copia_de_seguranca(window):
    window['-Mensagens-'].update('Criando cópia de segurança...')
    
    p.hotkey('ctrl', 's')
    wait_img('gravar_copia.png', conf=0.9)
    p.hotkey('alt', 'a')
    
    while not find_img('gravacao_concluida.png', conf=0.9):
        if find_img('copia_ja_existe.png', conf=0.9):
            click_img('copia_ja_existe.png', conf=0.9)
            p.hotkey('alt', 's')
        
    p.hotkey('alt', 'n')
    while find_img('gravacao_concluida.png', conf=0.9):
        time.sleep(1)
    
    p.hotkey('alt', 'f4')
    wait_img('confirmar_saida.png', conf=0.9)
    p.hotkey('alt', 's')
    while find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)
    
    
def move_arquivo_usado(pasta_arquivos, pasta, arquivo, verifica=''):
    try:
        for arq in os.listdir(verifica):
            if arq == arquivo:
                os.remove(os.path.join(verifica, arq))
                break
    except:
        pass
    
    os.makedirs(pasta, exist_ok=True)

    shutil.move(os.path.join(pasta_arquivos, arquivo), os.path.join(pasta, arquivo))


def move_arquivo_log(pasta_log, arquivo):
    ano = int(datetime.now().year)
    arquivo_xml = arquivo.replace('.txt', '.xml')
    arquivo_log = arquivo.replace('.txt', '.log')
    pasta_log_original = r'C:\Arquivos de Programas RFB\Dirf' + str(ano) + r'\db\temp'
    os.makedirs(pasta_log, exist_ok=True)
    
    shutil.copy(os.path.join(pasta_log_original, arquivo_log), os.path.join(pasta_log, arquivo_log))
    shutil.copy(os.path.join(pasta_log_original, arquivo_xml), os.path.join(pasta_log, arquivo_xml))


def deletar_relatorios_antigos(cnpj, arq, pasta_relatorios):
    for relatorio in pasta_relatorios:
        if relatorio == f'{cnpj} - {arq} - Relatório de erro.pdf':
            try:
                os.remove(os.path.join(pasta_relatorios, f'{cnpj} - {arq} - Relatório de erro.pdf'))
            except:
                pass


def fechar_dirf():
    p.hotkey('alt', 'f4')
    wait_img('confirmar_saida.png', conf=0.9)
    p.hotkey('alt', 's')
    while find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)


def transmitir_dirf(event, certificado, cnpj, arquivo, empresa, pasta_arquivos_importados, pasta_arquivos_nao_importados, pasta_arquivos_importados_com_erros_avisos, nome, pasta_recibos, pasta_relatorios_gravacao):
    print('>>> Trasmitindo arquivo')
    p.hotkey('ctrl', 'g')
    timer = 0
    while not find_img('tela_selecao.png', conf=0.9):
        p.hotkey('ctrl', 'g')
        time.sleep(1)
        timer += 1
        if timer > 5:
            abrir_arquivo(event, cnpj, arquivo, empresa, pasta_arquivos_importados, pasta_arquivos_nao_importados, pasta_arquivos_importados, pasta_arquivos_importados_com_erros_avisos, '', '')
            timer = 0
        
    p.hotkey('alt', 'o')
    
    wait_img('tela_gravacao.png', conf=0.9)
    p.hotkey('alt', 'a')
    
    while not find_img('local_transmissao.png', conf=0.9):
        if find_img('avancar_transmissao.png', conf=0.9):
            p.hotkey('alt', 'a')
        else:
            if find_img('erro_gravar.png', conf=0.9):
                p.hotkey('alt', 'r')
                resultado, mensagem = imprimir_relatorio(event, empresa, pasta_relatorios_gravacao)
                p.hotkey('alt', 'f4')
                p.hotkey('alt', 'c')
                return 'erro', f'Erro ao gravar arquivo, relatório salvo{mensagem}'

    click_position_img('local_transmissao.png', '+', pixels_x=120, conf=0.9)
    time.sleep(0.5)
    
    p.write('SP')
    time.sleep(0.5)
    
    p.press('enter')
    time.sleep(0.5)
    
    p.hotkey('alt', 'a')
    
    while not find_img('deseja_transmitir.png', conf=0.99):
        if find_img('regravar.png', conf=0.9):
            p.hotkey('alt', 's')
        if find_img('ja_existe_recibo.png', conf=0.9):
            p.hotkey('alt', 'c')
            return 'erro', 'Já existe recibo de entrega, declaração já transmitida'
    
    click_img('transmitir_sim.png', conf=0.99)
    
    if certificado == 'Sim':
        if find_img('com_certificado.png', conf=0.99):
            click_img('com_certificado.png', conf=0.99)
    else:
        if find_img('sem_certificado.png', conf=0.99):
            click_img('sem_certificado.png', conf=0.99)
    time.sleep(0.2)
    
    p.hotkey('alt', 'a')
    
    if certificado == 'Sim':
        wait_img('rpem_contabil.png', conf=0.9)
        click_img('rpem_contabil.png', conf=0.9)
        time.sleep(0.5)
        
        wait_img('assinar.png', conf=0.9)
        p.hotkey('alt', 'a')
    
    wait_img('resultado_transmissao.png', conf=0.9)
    
    while not find_img('transmitido.png', conf=0.99):
        if find_img('declarante_baixado.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O CNPJ do declarante consta com situação cadastral BAIXADA no cadastro da Receita Federal do Brasil em data anterior ao ano-calendário a que se refere a DIRF'
        if find_img('inscricao_posterior.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O CNPJ do declarante consta com cadastro da Receita Federal do Brasil com data de inscrição posterior ao ano-calendário a que se refere a DIRF'
        if find_img('consta_dirf_extincao.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Ano calendário anterior consta DIRF de extinção, transmita uma DIRF retificadora sem marcar extinção'
        if find_img('ja_entregue.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Já foi entregue declaração original para o CNPJ'
        if find_img('impossivel_entregar_dirf.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O declarante consta no cadastro da Receita Federal do Brasil com natureza jurídica impeditiva de entrega de DIRF'
        if find_img('procuracao_cancelada.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Procuração eletrônica cadastrada para o detentor do certificado digital apresentado foi cancelada'
        if find_img('procuracao_rejeitada.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Procuração eletrônica cadastrada para o detentor do certificado digital apresentado foi rejeitada por unidade de atendimento da Secretaria da Receita Federal do Brasil'
        if find_img('procuracao_expirou.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Procuração eletrônica cadastrada para o detentor do certificado digital apresentado expirou'
        if find_img('cpf_diferente.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O CPF informado na declaração como responsável pelo CNPJ é diferente do que consta no cadastro da RFB'
        if find_img('nao_tem_procuracao.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'Não existe procuração eletrônica para o detentor do certificado digital apresentado'
        if find_img('cnpj_nulo.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O CNPJ do plano privado de Assistência à Saúde consta com situação cadastral BAIXADA ou NULA em data anterior do ano calendário que se refere a DIRF'
        if find_img('nao_pode_p_juridica.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', 'O declarante consta no cadastro da Receita Federal do Brasil com natureza jurídica impedida de entrega da DIRF como Pessoa Jurídica. A DIRF deve ser apresentada como declarante Pessoa Física'
        if find_img('cadastro_baixado.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro', ('O CNPJ do declarante consta no cadastro da Receita Federal do Brasil com a situação BAIXADA no ano calendário '
                            'que se refere a DIRF e não é uma declaração de extinção')
        
        if find_img('erro_servidor.png', conf=0.9):
            p.hotkey('alt', 'n')
            return 'erro_servidor', ''
        
    if find_img('imprimir_recibo.png', conf=0.99):
        click_img('imprimir_recibo.png', conf=0.99)
    time.sleep(0.2)
    
    p.hotkey('alt', 'n')
    
    wait_img('servico_impressao.png', conf=0.9)
    
    if not find_img('print_pdf.png', conf=0.9):
        click_position_img('nome_impressora.png', '+', pixels_x=85, conf=0.9)
    
    wait_img('seleciona_print_pdf.png', conf=0.9)
    click_img('seleciona_print_pdf.png', conf=0.9)
    time.sleep(0.5)
    
    click_img('imprimir_recibo_entrega.png', conf=0.9)
    
    imprimir_recibo(cnpj, arquivo, nome, pasta_recibos)
    if find_img('erro_impressora.png', conf=0.9):
        click_img('erro_impressora.png', conf=0.9)
        time.sleep(0.2)
        p.press('enter')
        while find_img('erro_impressora.png', conf=0.9):
            time.sleep(1)
        if find_img('impressao_cancelada.png', conf=0.9):
            click_img('impressao_cancelada.png', conf=0.9)
            time.sleep(0.2)
            p.press('enter')
            time.sleep(0.2)
        p.hotkey('alt', 'n')
        return 'ok', 'Arquivo transmitido, Erro ao salvar o recibo de entrega'
    
    return 'ok', 'Arquivo transmitido'
    
    
def imprimir_recibo(cnpj, arquivo, nome, pasta_recibos):
    wait_img('salvar_como_recibo.png', conf=0.9)
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    pyperclip.copy(f'Recibo de entrega DIRF - {cnpj} - {arquivo.replace(".txt", "")} - {nome.replace("/", "")}.pdf')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy(pasta_recibos)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    print('✔ Arquivo transmitido')
    time.sleep(5)


def deletar_arquivo():
    p.hotkey('alt', 'd')
    wait_img('excluir_arquivo.png', conf=0.9)
    click_img('excluir_arquivo.png', conf=0.9)

    wait_img('tela_excluir.png', conf=0.9)
    click_img('select.png', conf=0.9)
    
    wait_img('ok.png', conf=0.9)
    p.hotkey('alt', 'o')
    
    wait_img('deseja_excluir.png', conf=0.9)
    p.hotkey('alt', 's')


@barra_de_status
def run(window, event):
    pasta_arquivos_importados = os.path.join(pasta_arquivos, 'Arquivos Importados')
    pasta_arquivos_nao_importados = os.path.join(pasta_arquivos, 'Arquivos não importados')
    pasta_arquivos_transmitidos = os.path.join(pasta_arquivos, 'Arquivos Transmitidos')
    pasta_arquivos_nao_transmitidos = os.path.join(pasta_arquivos, 'Arquivos não transmitidos')
    pasta_arquivos_importados_com_erros_avisos = os.path.join(pasta_arquivos, 'Arquivos importados com erros ou avisos')
    pasta_log = os.path.join(pasta_arquivos, 'Log de erros')
    pasta_relatorios = os.path.join(pasta_arquivos, 'Relatórios de erros e avisos')
    pasta_relatorios_gravacao = os.path.join(pasta_arquivos, 'Relatórios de erros ao gravar arquivo')
    pasta_recibos = os.path.join(pasta_arquivos, 'Recibos de entrega')
    andamentos = 'Importar Arquivos DIRF'
    empresas = cria_dados(window, planilha_de_dados, pasta_arquivos)
    window['-Mensagens-'].update('Aguardando programa DIRF')
    
    total_empresas = empresas[:]
    
    abrir_dirf(event)
    
    window['-titulo-'].update('Rotina automática, não interfira.')
    for count, empresa in enumerate(empresas[:], start=1):
        # printa o indice da empresa que está sendo executada
        
        window['-Mensagens-'].update(f'{str(count - 1)} de {str(len(total_empresas))} | {str((len(total_empresas)) - count + 1)} Restantes')
        
        cnpj, arquivo, nome = empresa
        print('\n')
        print(arquivo)
        while True:
            resultado = abrir_arquivo(event, cnpj, arquivo, empresa, pasta_arquivos, pasta_arquivos_importados, pasta_arquivos_nao_importados, pasta_arquivos_importados_com_erros_avisos, pasta_relatorios, pasta_log)
            if resultado == 'erro critico':
                p.hotkey('alt', 'f4')
                wait_img('confirmar_saida.png', conf=0.9)
                p.hotkey('alt', 's')
                while find_img('tela_inicial.png', conf=0.9):
                    time.sleep(1)
                os.startfile('C:\Arquivos de Programas RFB\Dirf2024\pgdDirf.exe')
                abrir_dirf(event)
                
            elif resultado != '':
                if rotina == 'Importar arquivos':
                    escreve_relatorio_csv(f'{cnpj};{arquivo};{resultado}', nome=andamentos, local=pasta_arquivos)
                break
                
        if event == '-encerrar-':
            return
        
        transmissao = ''
        if rotina == 'Importar e transmitir arquivos':
            andamentos = 'Transmitir Arquivos DIRF'
            if not re.compile(r'Arquivo não importado').search(resultado):
                abrir_dirf(event)
                while True:
                    if resultado == 'Arquivo importado com sucesso;Declaração com erros ou avisos, relatório salvo':
                        situacao, transmissao = transmitir_dirf(event, certificado, cnpj, arquivo, empresa, pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_nao_importados, pasta_arquivos_importados_com_erros_avisos, nome, pasta_recibos, pasta_relatorios_gravacao)
                    else:
                        situacao, transmissao = transmitir_dirf(event, certificado, cnpj, arquivo, empresa, pasta_arquivos_importados, pasta_arquivos_nao_importados, pasta_arquivos_importados_com_erros_avisos, nome, pasta_recibos, pasta_relatorios_gravacao)
                        
                    if situacao != 'erro_servidor':
                        break
                if resultado == 'Arquivo importado com sucesso;Declaração com erros ou avisos, relatório salvo':
                    if situacao == 'ok':
                        move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_transmitidos, arquivo)
                    if situacao == 'erro':
                        move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_nao_transmitidos, arquivo)
                elif resultado == 'Arquivo importado com sucesso, Declaração já importada;Declaração com erros ou avisos, relatório salvo':
                    if situacao == 'ok':
                        move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_transmitidos, arquivo)
                    if situacao == 'erro':
                        move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_nao_transmitidos, arquivo)
                
                else:
                    try:
                        if situacao == 'ok':
                            move_arquivo_usado(pasta_arquivos_importados, pasta_arquivos_transmitidos, arquivo)
                        if situacao == 'erro':
                            move_arquivo_usado(pasta_arquivos_importados, pasta_arquivos_nao_transmitidos, arquivo)
                    except:
                        if situacao == 'ok':
                            move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_transmitidos, arquivo)
                        if situacao == 'erro':
                            move_arquivo_usado(pasta_arquivos_importados_com_erros_avisos, pasta_arquivos_nao_transmitidos, arquivo)
                            
                if transmissao == 'Arquivo transmitido, Erro ao salvar o recibo de entrega':
                    fechar_dirf()
                    os.startfile('C:\Arquivos de Programas RFB\Dirf2024\pgdDirf.exe')
                    abrir_dirf(event)
                    
                abrir_dirf(event)
                deletar_arquivo()
            else:
                transmissao = ';Não transmitido'
            abrir_dirf(event)
            
        escreve_relatorio_csv(f'{cnpj};{arquivo};{resultado};{transmissao}', nome=andamentos, local=pasta_arquivos)
        
        if count % 200 == 0:
            # a cada 200 arquivos, faz um backup
            fechar_dirf()
            window['-Mensagens-'].update('Reiniciando programa...')
            os.startfile('C:\Arquivos de Programas RFB\Dirf2024\pgdDirf.exe')
            abrir_dirf(event)
            
    fechar_dirf()
    os.startfile('C:\Arquivos de Programas RFB\Dirf2024\pgdDirf.exe')
        

if __name__ == '__main__':
    rotina = p.confirm(buttons=['Importar arquivos', 'Importar e transmitir arquivos'])
    if rotina == 'Importar e transmitir arquivos':
        certificado = p.confirm(text='Transmitir com Certificado Digital?', buttons=['Sim', 'Não'])
    else:
        certificado = 'Sim'
    
    empresas = None
    while True:
        pasta_arquivos = ask_for_dir()
        if not pasta_arquivos:
            break
            
        tem_arquivo = False
        for arq in os.listdir(pasta_arquivos):
            if arq.endswith('.txt'):
                tem_arquivo = True
        
        if tem_arquivo:
            break
        p.alert(text='Essa pasta não contem arquivos ".txt"')
        
    if pasta_arquivos:
        planilha_de_dados = p.confirm(buttons=['Criar planilha com os arquivos na pasta selecionada', 'Selecionar planilha já existente'])
        run()
        
    p.alert(text='Execução finalizada.')
    