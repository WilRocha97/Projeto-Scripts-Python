# -*- coding: utf-8 -*-
import json, io, os, traceback, PySimpleGUI as sg
from threading import Thread
from pyautogui import confirm, alert
from PIL import Image, ImageDraw
from PySimpleGUI import BUTTON_TYPE_READ_FORM, FILE_TYPES_ALL_FILES, theme_background_color, theme_button_color, BUTTON_TYPE_BROWSE_FILE, BUTTON_TYPE_BROWSE_FOLDER
from base64 import b64encode

import comum, rotinas


icone = 'Assets/auto-flash.ico'
dados_modo = 'Assets/modo.txt'
controle_rotinas = 'Log/Controle.txt'
controle_botoes = 'Log/Buttons.txt'
dados_elementos = 'Log/window_values.json'


# Função para salvar os valores dos elementos em um arquivo JSON
def save_values(values, filename=dados_elementos):
    with open(filename, 'w', encoding='latin-1') as f:
        json.dump(values, f)

# Função para carregar os valores dos elementos de um arquivo JSON
def load_values(filename=dados_elementos):
    if os.path.exists(filename):
        with open(filename, 'r', encoding='latin-1') as f:
            return json.load(f)
    return {}


# Define o ícone global da aplicação
sg.set_global_icon(icone)
if __name__ == '__main__':
    # limpa o arquivo que salva o estado dos elementos preenchidos e o estado anterior do botão de abrir os resultados
    with open(dados_elementos, 'w', encoding='utf-8') as f:
        f.write('')
    with open(controle_botoes, 'w', encoding='utf-8') as f:
        f.write('')
    
    sg.LOOK_AND_FEEL_TABLE['tema_claro'] = {'BACKGROUND': '#F8F8F8',
                                            'TEXT': '#000000',
                                            'INPUT': '#F8F8F8',
                                            'TEXT_INPUT': '#000000',
                                            'SCROLL': '#F8F8F8',
                                            'BUTTON': ('#000000', '#F8F8F8'),
                                            'PROGRESS': ('#fca400', '#D7D7D7'),
                                            'BORDER': 0,
                                            'SLIDER_DEPTH': 0,
                                            'PROGRESS_DEPTH': 0}
    sg.LOOK_AND_FEEL_TABLE['tema_escuro'] = {'BACKGROUND': '#000000',
                                             'TEXT': '#F8F8F8',
                                             'INPUT': '#000000',
                                             'TEXT_INPUT': '#F8F8F8',
                                             'SCROLL': '#000000',
                                             'BUTTON': ('#F8F8F8', '#000000'),
                                             'PROGRESS': ('#fca400', '#2A2A2A'),
                                             'BORDER': 0,
                                             'SLIDER_DEPTH': 0,
                                             'PROGRESS_DEPTH': 0}
    
    # define o tema baseado no tema que estava selecionado na última vêz que o programa foi fechado
    f = open(dados_modo, 'r', encoding='utf-8')
    modo = f.read()
    sg.theme(f'tema_{modo}')
    
    tamanho_padrao = (680, 550)
    
    def RoundedButton(button_text=' ', corner_radius=0.1, button_type=BUTTON_TYPE_READ_FORM, target=(None, None), tooltip=None, file_types=FILE_TYPES_ALL_FILES, initial_folder=None, default_extension='',
                      disabled=False, change_submits=False, enable_events=False, image_size=(None, None), image_subsample=None, border_width=None, size=(None, None), auto_size_button=None,
                      button_color=None, disabled_button_color=None, highlight_colors=None, mouseover_colors=(None, None), use_ttk_buttons=None, font=None, bind_return_key=False, focus=False,
                      pad=None, key=None, right_click_menu=None, expand_x=False, expand_y=False, visible=True, metadata=None):
        if None in size:
            multi = 5
            size = (((len(button_text) if size[0] is None else size[0]) * 5 + 20) * multi,
                    20 * multi if size[1] is None else size[1])
        if button_color is None:
            button_color = theme_button_color()
        btn_img = Image.new('RGBA', size, (0, 0, 0, 0))
        corner_radius = int(corner_radius / 2 * min(size))
        poly_coords = (
            (corner_radius, 0),
            (size[0] - corner_radius, 0),
            (size[0], corner_radius),
            (size[0], size[1] - corner_radius),
            (size[0] - corner_radius, size[1]),
            (corner_radius, size[1]),
            (0, size[1] - corner_radius),
            (0, corner_radius),
        )
        pie_coords = [
            [(size[0] - corner_radius * 2, size[1] - corner_radius * 2, size[0], size[1]),
             [0, 90]],
            [(0, size[1] - corner_radius * 2, corner_radius * 2, size[1]), [90, 180]],
            [(0, 0, corner_radius * 2, corner_radius * 2), [180, 270]],
            [(size[0] - corner_radius * 2, 0, size[0], corner_radius * 2), [270, 360]],
        ]
        brush = ImageDraw.Draw(btn_img)
        brush.polygon(poly_coords, button_color[1])
        for coord in pie_coords:
            brush.pieslice(coord[0], coord[1][0], coord[1][1], button_color[1])
        data = io.BytesIO()
        btn_img.thumbnail((size[0] // 3, size[1] // 3), resample=Image.LANCZOS)
        btn_img.save(data, format='png', quality=95)
        btn_img = b64encode(data.getvalue())
        return sg.Button(button_text=button_text, button_type=button_type, target=target, tooltip=tooltip,
                         file_types=file_types, initial_folder=initial_folder, default_extension=default_extension,
                         disabled=disabled, change_submits=change_submits, enable_events=enable_events,
                         image_data=btn_img, image_size=image_size,
                         image_subsample=image_subsample, border_width=border_width, size=size,
                         auto_size_button=auto_size_button, button_color=(button_color[0], theme_background_color()),
                         disabled_button_color=disabled_button_color, highlight_colors=highlight_colors,
                         mouseover_colors=mouseover_colors, use_ttk_buttons=use_ttk_buttons, font=font,
                         bind_return_key=bind_return_key, focus=focus, pad=pad, key=key, right_click_menu=right_click_menu,
                         expand_x=expand_x, expand_y=expand_y, visible=visible, metadata=metadata)
    
    def janela_principal():
        def cria_layout():
            rotinas = ['Consulta Certidão Negativa de Débitos Tributários Não Inscritos', 'Consulta Débitos Municipais Jundiaí', 'Consulta Débitos Municipais Valinhos',
                       'Consulta Débitos Municipais Vinhedo', 'Consulta Débitos Estaduais - Situação do Contribuinte', 'Consulta Divida Ativa', 'Consulta Pendências SIGISSWEB Valinhos',]
            
            return [[sg.Button('AJUDA'),
                     sg.Button('SOBRE'),
                     sg.Button('LOG DO SISTEMA', key='-log_sistema-', disabled=True),
                     sg.Button('CONFIGURAÇÕES', key='-config-'),
                     sg.Text('', expand_x=True),
                     sg.Button('☀', key='-tema-', border_width=0)],
                    [sg.Text('')],
                    [sg.Text('')],

                    [sg.Combo(rotinas, expand_x=True, enable_events=True, readonly=True, key='-dropdown-', default_value='Selecione a rotina que será executada')],
                    [sg.Text('')],

                    [sg.pin(sg.Text(text='Selecione a pasta para salvar os resultados:', key='-pasta_resultados_texto-', visible=False)),
                     sg.pin(RoundedButton('Selecione a pasta para salvar os resultados:', key='-pasta_resultados-', corner_radius=0.8, button_color=(None, '#848484'),
                                          size=(900, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-output_dir-')),
                     sg.pin(sg.InputText(key='-output_dir-', text_color='#fca400', expand_x=True))],

                    [sg.pin(sg.Text(text='Selecione a planilha de dados:', key='-planilha_dados_texto-', visible=False)),
                     sg.pin(RoundedButton('Selecione a planilha de dados:', key='-planilha_dados-', corner_radius=0.8, button_color=(None, '#848484'), button_type=BUTTON_TYPE_BROWSE_FILE,
                                          size=(650, 100), file_types=(('Planilhas Excel', ['*.xlsx', '*.xls']),), target='-input_dados-')),
                     sg.pin(sg.InputText(key='-input_dados-', text_color='#fca400', expand_x=True))],
                    [sg.Text('')],

                    [sg.Text('Utilizar empresas com o código acima de 20.000')],
                    [sg.Checkbox(key='-codigo_20000_sim-', text='Sim', enable_events=True),
                     sg.Checkbox(key='-codigo_20000_nao-', text='Não', enable_events=True, default=True),
                     sg.Checkbox(key='-codigo_20000-', text='Apenas acima do 20.000', enable_events=True)],
                    [sg.Text('', expand_y=True)],

                    [sg.Text('', key='-Mensagens_2-')],
                    [sg.Text('', key='-Mensagens-')],
                    [sg.Text(size=6, text='', key='-Progresso_texto-'),
                     sg.ProgressBar(max_value=0, orientation='h', size=(5, 5), key='-progressbar-', expand_x=True, visible=False)],
                    [sg.pin(RoundedButton('INICIAR', key='-iniciar-', corner_radius=0.8, button_color=(None, '#fca400'))),
                     sg.pin(RoundedButton('ENCERRAR', key='-encerrar-', corner_radius=0.8, button_color=(None, '#fca400'), visible=False)),
                     sg.pin(RoundedButton('ABRIR RESULTADOS', key='-abrir_resultados-', corner_radius=0.8, button_color=(None, '#fca400'), visible=False))]
            ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Consultas ADM', cria_layout(), resizable=True, finalize=True, margins=(30, 30))
    
    def janela_configura():  # layout da janela do menu principal
        def cria_layout_configura():
            return [
                [sg.Text('Insira a nova chave de acesso para a API Anticaptcha')],
                [sg.InputText(key='-input_chave_api-', size=90, password_char='*', default_text='', border_width=1)],
                [sg.Button('APLICAR', key='-confirma_conf-', border_width=0),
                 sg.Button('CANCELAR', key='-cancela_conf-', border_width=0), ]
            ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Configurações', cria_layout_configura(), finalize=True, modal=True, margins=(30, 30))
    
    
    def run_script_thread():
        if rotina == 'Selecione a rotina que será executada':
            alert('❗ Por favor informe uma rotina para executar')
            return
        
        # verifica se os inputs foram preenchidos corretamente
        for elemento in [(rotina, 'Por favor informe uma rotina para executar'), (pasta_final, 'Por favor informe um diretório para salvar os andamentos.'),
                         (planilha_dados, 'Por favor informe uma rotina para executar')]:
            if not elemento[0]:
                alert(f'❗ {elemento[1]}')
                return
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-iniciar-', False), ('-encerrar-', True), ('-abrir_resultados-', True), ('-planilha_dados-', False), ('-pasta_resultados-', False),
                    ('-planilha_dados_texto-', True), ('-pasta_resultados_texto-', True), ('-progressbar-', True)]:
            window_principal[key[0]].update(visible=key[1])
        
        for key in [('-tema-', True), ('-dropdown-', True), ('-codigo_20000_sim-', True),
                    ('-codigo_20000_nao-', True), ('-codigo_20000-', True), ('-config-', True)]:
            window_principal[key[0]].update(disabled=key[1])
        
        # controle para saber se os botões estavam visíveis ao trocar o tema da janela
        with open(controle_botoes, 'w', encoding='utf-8') as f:
            f.write('visible')
        
        # apaga qualquer mensagem na interface
        window_principal['-Mensagens-'].update('')
        window_principal['-Mensagens_2-'].update('')

        try:
            # Chama a função que executa o script
            rotinas.executa_rotina(window_principal, codigo_20000, planilha_dados, pasta_final, rotina)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            comum.escreve_doc(f'Traceback: {traceback_str}\n\n'
                              f'Erro: {erro}')
            window_principal['-log_sistema-'].update(disabled=False)
            alert('❌ Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
        
        window_principal['-progressbar-'].update_bar(0)
        window_principal['-Progresso_texto-'].update('')
        window_principal['-Mensagens-'].update('')
        window_principal['-Mensagens_2-'].update('')
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-progressbar-', False), ('-encerrar-', False), ('-iniciar-', True), ('-planilha_dados-', True), ('-pasta_resultados-', True),
                    ('-planilha_dados_texto-', False), ('-pasta_resultados_texto-', False)]:
            window_principal[key[0]].update(visible=key[1])
        
        for key in [('-tema-', False), ('-dropdown-', False), ('-codigo_20000_sim-', False),
                    ('-codigo_20000_nao-', False), ('-codigo_20000-', False), ('-config-', False)]:
            window_principal[key[0]].update(disabled=key[1])
        
    # inicia as variáveis das janelas
    window_principal, window_configura = janela_principal(), None
    # Definindo o tamanho mínimo da janela
    window_principal.set_min_size(tamanho_padrao)
    
    codigo_20000 = '-codigo_20000_nao-'
    while True:
        # configura o tema conforme o tema que estava quando o programa foi fechado da última ves
        f = open(dados_modo, 'r', encoding='utf-8')
        modo = f.read()
        if modo == 'claro':
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-planilha_dados-', '-pasta_resultados-']:
                window_principal[key].update(button_color=('#F8F8F8', None))
        else:
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-planilha_dados-', '-pasta_resultados-']:
                window_principal[key].update(button_color=('#000000', None))
            
        # captura o evento e os valores armazenados na interface
        window, event, values = sg.read_all_windows()
        
        # recebe os valores dos inputs
        try:
            pasta_final = values['-output_dir-']
        except:
            pasta_final = None
        try:
            rotina = values['-dropdown-']
        except:
            rotina = None
        try:
            planilha_dados = values['-input_dados-']
        except:
            planilha_dados = None
        
        # salva o estado da interface
        save_values(values)
        
        checkboxes = ['-codigo_20000_sim-', '-codigo_20000_nao-', '-codigo_20000-']
        if event in ('-codigo_20000_sim-', '-codigo_20000_nao-', '-codigo_20000-'):
            for checkbox in checkboxes:
                if checkbox != event:
                    window[checkbox].update(value=False)
                else:
                    codigo_20000 = checkbox
        
        if event == sg.WIN_CLOSED:
            if window == window_configura:  # if closing win 2, mark as closed
                window_configura = None
            elif window == window_principal:  # if closing win 1, exit program
                break
        
        # troca o tema
        elif event == '-tema-':
            f = open(dados_modo, 'r', encoding='utf-8')
            modo = f.read()
            
            if modo == 'claro':
                sg.theme('tema_escuro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('escuro')
            else:
                sg.theme('tema_claro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('claro')
    
            window_principal.close()  # Fecha a janela atual
            # inicia as variáveis das janelas
            window_principal, window_configura = janela_principal(), None
            # Definindo o tamanho mínimo da janela
            window_principal.set_min_size(tamanho_padrao)
            # retorna os elementos preenchidos
            values = load_values()
            for key, value in values.items():
                if key in window_principal.AllKeysDict:
                    try:
                        window_principal[key].update(value)
                    except Exception as e:
                        print(f"Erro ao atualizar o elemento {key}: {e}")
            
            window_principal['-planilha_dados-'].update('Selecione a planilha de dados:')
            window_principal['-pasta_resultados-'].update('Selecione a pasta para salvar os resultados:')
            try:
                pasta_final = values['-output_dir-']
            except:
                pasta_final = None
            try:
                rotina = values['-dropdown-']
            except:
                rotina = None
            try:
                planilha_dados = values['-input_dados-']
            except:
                planilha_dados = None
            
            # recupera o estado do botão de abrir resultados
            cb = open(controle_botoes, 'r', encoding='utf-8').read()
            if cb == 'visible':
                window_principal['-abrir_resultados-'].update(visible=True)
            
            
        elif event == '-config-':
            window_configura = janela_configura()
            
            while True:
                # captura o evento e os valores armazenados na interface
                event, values = window_configura.read()
                if event == '-confirma_conf-':
                    confirma = confirm(text='As alterações serão aplicadas, deseja continuar?')
                    nova_chave = values['-input_chave_api-']
                    if confirma == 'OK':
                        with open(comum.dados_anticaptcha, 'w', encoding='utf-8') as f:
                            f.write(nova_chave)
                        break
                if event == '-cancela_conf-':
                    confirma = confirm(text='As alterações serão perdidas, deseja continuar?')
                    if confirma == 'OK':
                        break
                if event == sg.WIN_CLOSED:
                    confirma = confirm(text='As alterações serão perdidas, deseja continuar?')
                    if confirma == 'OK':
                        break
                    window_configura = janela_configura()
            
            window_configura.close()
        
        elif event == '-log_sistema-':
            os.startfile('Log')
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Cria E-book DIRPF.pdf')
        
        elif event == 'Sobre':
            os.startfile('Sobre.pdf')
        
        elif event == '-iniciar-':
            with open(controle_rotinas, 'w', encoding='utf-8') as f:
                f.write('START')
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
            
        elif event == '-encerrar-':
            window_principal['-Mensagens_2-'].update('')
            window_principal['-Mensagens-'].update('Encerrando, aguarde...')
            with open(controle_rotinas, 'w', encoding='utf-8') as f:
                f.write('STOP')

        elif event == '-abrir_resultados-':
            os.makedirs(os.path.join(pasta_final, rotina), exist_ok=True)
            os.startfile(os.path.join(pasta_final, rotina))
        
    window_principal.close()
