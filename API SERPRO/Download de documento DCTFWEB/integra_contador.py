# -*- coding: utf-8 -*-
import traceback, requests_pkcs12, atexit, sys, re, os, requests, time, base64, json, io, chardet, OpenSSL.crypto, contextlib, tempfile, pyautogui as p, PySimpleGUI as sg
from PySimpleGUI import BUTTON_TYPE_READ_FORM, FILE_TYPES_ALL_FILES, theme_background_color, theme_button_color, BUTTON_TYPE_BROWSE_FILE, BUTTON_TYPE_BROWSE_FOLDER
from PIL import Image, ImageDraw
from base64 import b64encode
from threading import Thread
from pathlib import Path
from functools import wraps

dados = "info\\info.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.readline()
input_consumer = user.split('/')

icone = 'Assets/AS_icon.ico'
dados_modo = 'Assets/modo.txt'
controle_botoes = 'Log/Buttons.txt'
dados_elementos = 'Log/window_values.json'


def create_lock_file(lock_file_path):
    try:
        # Tente criar o arquivo de trava
        with open(lock_file_path, 'x') as lock_file:
            lock_file.write(str(os.getpid()))
        return True
    except FileExistsError:
        # O arquivo de trava já existe, indicando que outra instância está em execução
        return False


def remove_lock_file(lock_file_path):
    try:
        os.remove(lock_file_path)
    except FileNotFoundError:
        pass
    
    
# abre a lista de dados da empresa em .csv
def open_lista_dados(input_excel, encode='latin-1'):
    try:
        with open(input_excel, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        p.alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


# escreve os andamentos das requisições em um .csv
def escreve_relatorio_csv(texto, nome='resumo', local='Desktop', end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome}-auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


# escreve arquivos .txt com as respostas da API
def escreve_doc(texto, local='Log', nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)

    f.write(str(texto))
    f.close()


# converte as chaves de acesso do site da API em base64 para fazer a requisição das chaves de acesso para o serviço
def converter_base64(usuario):
    # converte as credenciais para base64
    return base64.b64encode(usuario.encode("utf8")).decode("utf8")


# solicita as chaves de acesso para o serviço
def solicita_token(usuario_b64, input_certificado, senha, output_dir):
    headers = {
        "Authorization": "Basic " + usuario_b64,
        "role-type": "TERCEIROS",
        "content-type": "application/x-www-form-urlencoded"
    }
    body = {'grant_type': 'client_credentials'}
    pagina = requests_pkcs12.post('https://autenticacao.sapi.serpro.gov.br/authenticate',
                                  data=body,
                                  headers=headers,
                                  verify=True,
                                  pkcs12_filename=input_certificado,
                                  pkcs12_password=senha,
                                  timeout=30)
    
    resposta_string_json = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    resposta = pagina.json()
    
    # anota as respostas para tratar possíveis erros
    escreve_doc(pagina.status_code, nome='status_code', local=output_dir)
    escreve_doc(resposta, nome='resposta_jason', local=output_dir)
    escreve_doc(resposta_string_json, nome='string_json', local=output_dir)
    
    #
    # Output:
    #   200
    #   {
    #       "access_token": "05e658b6-212e-3f0e-b068-9d125a3ea479",
    #       "expires_in": 3032,
    #       "jwt_pucomex": null,
    #       "jwt_token": "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIzMzY4MzExMTxj...",
    #       "scope": "am_application_scope default",
    #       "token_type": "Bearer"
    #   }
    #
    try:
        return resposta['jwt_token'], resposta['access_token']
    except:
        try:
            return resposta['error'], resposta['message']
        except:
            return resposta['message'], resposta['description']
    

# solicita a guia de DCTF WEB na API
def solicita_dctf(output_dir, tipo, categoria, comp, cnpj_contratante, id_empresa, access_token, jwt_token):
    # verifica se a guia será para CPF ou CNPJ, se for CPF é código 1 e CNPJ é código 2
    if len(id_empresa) > 12:
        cod_consulta = 2
        tipo_categoria = 'GERAL'
    else:
        cod_consulta = 1
        tipo_categoria = 'PF'
        
    if tipo == '-guia_pagamento-':
        id_servico = "GERARGUIA31"
        url = 'https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir'
    else:
        id_servico = "CONSDECCOMPLETA33"
        url = 'https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Consultar'
    
      
    categoria_guias = ''
    if categoria == '-categoria_mensal-':
        mes = comp.split('/')[0]
        ano = comp.split('/')[1]
        categoria_guias = "{\"categoria\": \"" + tipo_categoria + "_MENSAL\",\"anoPA\":\"" + str(ano) + "\",\"mesPA\":\"" + str(mes) + "\"}"
    if categoria == '-categoria_13-':
        mes = '13'
        ano = comp
        categoria_guias = "{\"categoria\": \"" + tipo_categoria + "_13o_SALARIO\",\"anoPA\":\"" + str(ano) + "\"}"
    
    data = {
              "contratante": {
                "numero": str(cnpj_contratante),
                "tipo": 2
              },
              "autorPedidoDados": {
                "numero": str(cnpj_contratante),
                "tipo": 2
              },
              "contribuinte": {
                "numero": str(id_empresa),
                "tipo": int(cod_consulta)
              },
              "pedidoDados": {
                "idSistema": "DCTFWEB",
                "idServico": str(id_servico),
                "versaoSistema": "1.0",
                "dados": str(categoria_guias)
              }
            }
    
    headers = {'accept': 'text/plain',
               'Authorization': 'Bearer ' + access_token,
               'Content-Type': 'application/json',
               'jwt_token': jwt_token}
    
    tentativas = 0
    while True:
        pagina = requests.post(url, headers=headers, data=json.dumps(data))
        resposta = pagina.json()
        
        try:
            resposta['message']
        except:
            break
            
        tentativas += 1
        
        if tentativas >= 10:
            break
            
    resposta_string_json = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    
    # anota as respostas da API para tratar possíveis erros
    os.makedirs(output_dir, exist_ok=True)
    escreve_doc(resposta, nome='resposta_jason_guia', local=output_dir)
    escreve_doc(f'{id_empresa}\n{resposta_string_json}', nome='string_json_guia', local=output_dir)
    #
    # output
    # {
    #     "contratante":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "autorPedidoDados":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "contribuinte":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "pedidoDados":
    #         {
    #             "idSistema": "DCTFWEB",
    #             "idServico": "GERARGUIA31",
    #             "versaoSistema": "1.0",
    #             "dados": "{\"categoria\": \"GERAL_MENSAL\",\"anoPA\":\"2027\",\"mesPA\":\"11\"}"
    #         },
    #     "status": 200,
    #     "dados": "{\"PDFByteArrayBase64\":\"JVBERi0xLjUKJafj8fEKMi..."}",
    #     "mensagens":
    #         [
    #             {
    #                 "codigo": "Aviso-DCTFWEB-MG11",
    #                 "texto": "Emissor Guia Pagamento executado com sucesso."
    #             },
    #             {
    #                 "codigo": "Sucesso-DCTFWEB-MG00",
    #                 "texto": "Requisição efetuada com sucesso"
    #             }
    #         ]
    # }
    #
    try:
        escreve_doc(resposta['dados'], nome='pdf_base_64', local=output_dir)
        dados_pdf = json.loads(resposta["dados"])
        return mes, ano, dados_pdf["PDFByteArrayBase64"], resposta['mensagens'][0]['texto']
    except:
        try:
            return mes, ano, resposta['dados'], resposta['mensagens'][2]['texto']
        except:
            try:
                return mes, ano, resposta['dados'], resposta['mensagens'][0]['texto']
            except:
                return mes, ano, None, resposta['message']
            

# cria o PDF usando os bytes retornados da requisição na API
def cria_pdf(pdf_base64, output_dir, tipo_servico, id_empresa, nome_empresa, mes, ano):
    # limpa o nome da empresa para não dar erro no arquivo
    nome_empresa = nome_empresa.replace('/', '').replace(',', '')
    
    # verifica se a pasta para salvar o PDF existe, se não então cria
    e_dir_guias = os.path.join(output_dir, f'{tipo_servico.capitalize()} {mes} - {ano}')
    os.makedirs(e_dir_guias, exist_ok=True)
    
    # decodifica a base64 em bytes
    pdf_bytes = base64.b64decode(pdf_base64)
    # cria o PDF a partir dos bytes
    with open(os.path.join(e_dir_guias, f'DCTFWEB {tipo_servico} - {mes}-{ano} - {id_empresa} - {nome_empresa[:70]}.pdf'), "wb") as file:
        file.write(pdf_bytes)


def run(window, cnpj_contratante, usuario_b64, senha, tipo, categoria, competencia, input_certificado, input_excel, output_dir):
    if tipo == '-guia_pagamento-':
        tipo_servico = "guias"
    else:
        tipo_servico = "declarações"
    # cria a pasta onde serão salvos os resultados, no caso sempre será criada uma pasta diferente da criada anteriormente
    contador = 0
    while True:
        try:
            os.makedirs(os.path.join(output_dir, f'Download de {tipo_servico} DCTFWEB SERPRO'))
            output_dir = os.path.join(output_dir, f'Download de {tipo_servico} DCTFWEB SERPRO')
            break
        except:
            try:
                contador += 1
                os.makedirs(os.path.join(output_dir, f'Download de {tipo_servico} DCTFWEB SERPRO ({str(contador)})'))
                output_dir = os.path.join(output_dir, f'Download de {tipo_servico} DCTFWEB SERPRO ({str(contador)})')
                break
            except:
                pass
    local = os.path.join(output_dir, 'API_response')
    os.makedirs(local, exist_ok=True)
    
    for arq in os.listdir(local):
        os.remove(os.path.join(local, arq))
        
    if event == '-encerrar-' or event == sg.WIN_CLOSED:
        return
        
    # solicita os tokens para realizar a emissão das guias
    jwt_token, access_token = solicita_token(usuario_b64, input_certificado, senha, output_dir)
    
    if jwt_token == 'Unauthorized':
        p.alert(text='Consumer Secret ou Consumer Key inválido')
        return
    
    # abrir a planilha de dados
    empresas = open_lista_dados(input_excel)
    if not empresas:
        return False
    
    time.sleep(1)
    
    window['-Mensagens-'].update(f'Buscando guias (0 de {len(empresas)})')
    window.refresh()
    
    if event == '-encerrar-' or event == sg.WIN_CLOSED:
        try:
            os.rmdir(output_dir)
        except:
            pass
        return
    
    for count, empresa in enumerate(empresas, start=1):
        id_empresa, nome_empresa = empresa
        # solicita a guia de DCTF
        mes, ano, pdf_base64, mensagens = solicita_dctf(output_dir, tipo, categoria, competencia, cnpj_contratante, id_empresa, str(access_token), str(jwt_token))

        if re.compile(r'Acesso negado').search(mensagens):
            p.alert(text=mensagens)
            return
        
        # se não retornar o PDF não precisa da segunda mensagem
        if not pdf_base64:
            mensagen_2 = ''
            if mensagens == 'Runtime Error':
                mensagen_2 = 'Erro ao acessar a API'
                
        # se retornar o PDF
        else:
            try:
                # tenta converter a base64 em PDF e não precisa da segunda mensagem
                cria_pdf(pdf_base64, output_dir, tipo_servico, id_empresa, nome_empresa, mes, ano)
                mensagen_2 = ''
            # se não converter o PDF captura a segunda mensagem
            except Exception as e:
                mensagen_2 = f'Não gerou PDF {e}'
        
        # escreve os andamentos
        escreve_relatorio_csv(f'{id_empresa};{nome_empresa};{mensagens};{mensagen_2}', local=output_dir, nome=f'Andamentos DCTF-WEB {mes}-{ano}')
        
        # atualiza a barra de progresso
        window['-Mensagens-'].update(f'Buscando guias ({count} de {len(empresas)})')
        window['-progressbar-'].update_bar(count, max=int(len(empresas)))
        window['-Progresso_texto-'].update(str(round(float(count) / int(len(empresas)) * 100, 1)) + '%')
        window.refresh()
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            p.alert(text='Download encerrado.\n\n')
            return
        
    p.alert(text='Download finalizado!')


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
    # Especifique o caminho para o arquivo de trava
    lock_file_path = 'integra_contador.lock'
    
    # Verifique se outra instância está em execução
    if not create_lock_file(lock_file_path):
        p.alert(text="Outra instância já está em execução.")
        sys.exit(1)
    
    # Defina uma função para remover o arquivo de trava ao final da execução
    atexit.register(remove_lock_file, lock_file_path)
    
    categoria_key = ('GERAL_MENSAL', 'GERAL_13o_SALARIO')
    categoria_nome = ('Mensal', '13º')
    
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
    
    
    # sg.theme_previewer()
    # Layout da janela
    def cria_layout():
        layout = [
                [sg.Button('AJUDA', border_width=0),
                   sg.Button('LOG DO SISTEMA', border_width=0),
                   sg.Text('', expand_x=True),
                   sg.Button('', key='-tema-', font=("Arial", 11), border_width=0)],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('Informe o CNPJ do contratante do serviço SERPRO:'), sg.InputText(key='-input_cnpj_contratante-', size=14, default_text=input_consumer[0])],
                [sg.Text('Informe a Consumer Key:'), sg.InputText(key='-input_consumer_key-', size=30, password_char='*', default_text=input_consumer[1])],
                [sg.Text('Informe a Consumer Secret:'), sg.InputText(key='-input_consumer_secret-', size=30, password_char='*', default_text=input_consumer[2])],
                [sg.Text('Informe a senha do certificado digital:'), sg.InputText(key='-input_senha_certificado-', size=30, password_char='*')],
                [sg.Checkbox(key='-categoria_mensal-', text='Mensal', enable_events=True), sg.Checkbox(key='-categoria_13-', text='13º', enable_events=True)],
                [sg.Checkbox(key='-guia_pagamento-', text='Documento de arrecadação', enable_events=True), sg.Checkbox(key='-declaração-', text='Declaração Completa', enable_events=True)],
                [sg.Text('Informe a competência das guias. (Mensal = "00/0000") (13º = "0000"):'), sg.InputText(key='-input_competencia-', size=7)],
                [sg.Text('')],
                
                [sg.pin(sg.Text(text='Selecione o certificado digital', key='-abrir_texto-', visible=False)),
                 sg.pin(RoundedButton('Selecione o certificado digital', key='-abrir-', corner_radius=0.8, button_color=(None, '#848484'), button_type=BUTTON_TYPE_BROWSE_FILE,
                                      size=(650, 100), file_types=(('PFX files', '*.pfx'),), target='-input_certificado-')),
                 sg.pin(sg.InputText(expand_x=True, key='-input_certificado-', text_color='#fca400'), expand_x=True)],
                
                [sg.pin(sg.Text(text='Selecione um arquivo Excel com os dados dos clientes', key='-abrir1_texto-', visible=False)),
                 sg.pin(RoundedButton('Selecione um arquivo Excel com os dados dos clientes', key='-abrir1-', corner_radius=0.8, button_color=(None, '#848484'), button_type=BUTTON_TYPE_BROWSE_FILE,
                                      size=(1100, 100), file_types=(('Planilhas Excel', '*.csv'),), target='-input_excel-')),
                 sg.pin(sg.InputText(expand_x=True, key='-input_excel-', text_color='#fca400'), expand_x=True)],
                
                [sg.pin(sg.Text(text='Selecione um diretório para salvar os resultados (Servidor "Comum T:" é o padrão)', key='-abrir2_texto-', visible=False)),
                 sg.pin(RoundedButton('Selecione um diretório para salvar os resultados (Servidor "Comum T" é o padrão):', key='-abrir2-', corner_radius=0.8, button_color=(None, '#848484'),
                                      size=(1550, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-output_dir-')),
                 sg.pin(sg.InputText(expand_x=True, default_text='T:/ROBÔ/DCTF-WEB', key='-output_dir-', text_color='#fca400'), expand_x=True)],
                
                [sg.Text('', expand_y=True)],
                
                [sg.Text('', key='-Mensagens-')],
                [sg.Text(text='', key='-Progresso_texto-'),
                 sg.ProgressBar(max_value=0, orientation='h', size=(5, 5), expand_x=True, key='-progressbar-', visible=False)],
                [sg.pin(RoundedButton('INICIAR', key='-iniciar-', corner_radius=0.8, button_color=(None, '#fca400'))),
                 sg.pin(RoundedButton('ENCERRAR', key='-encerrar-', corner_radius=0.8, button_color=(None, '#fca400'), visible=False)),
                 sg.pin(RoundedButton('ABRIR RESULTADOS', key='-abrir_resultados-', corner_radius=0.8, button_color=(None, '#fca400'), visible=False))]
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Download de documentos DCTFWEB API SERPRO', layout, resizable=True, finalize=True, margins=(30, 30))
    
    
    def run_script_thread():
        # verifica se os inputs foram preenchidos corretamente
        for elemento in [(cnpj_contratante, 'Por favor informe o CNPJ do contratante da API SERPRO.'), (len(cnpj_contratante) == 14, 'Por favor informe um CNPJ válido.'),
                         (consumer_key, 'Por favor informe o consumerKey.'), (consumer_secret, 'Por favor informe o consumerSecret.'), (senha, 'Por favor informe a senha do certificado digital.'),
                         (categoria, 'Por favor informe se a categoria é Mensal ou 13º.'), (tipo, 'Por favor informe se o tipo de documento é o Documento de Arrecadação ou Declaração completa'),
                         (competencia, 'Por favor informe a competência das guias.'), (input_certificado, 'Por favor selecione um certificado digital.'),
                         (input_excel, 'Por favor selecione uma planilha de dados.')]:
            if not elemento[0]:
                p.alert(f'❗ {elemento[1]}')
                return
        
        if categoria == '-categoria_13-':
            if len(competencia) > 4:
                p.alert(text=f'Por favor insira apenas o ano referente a competência de 13º.')
                return
        else:
            if not re.compile(r'\d\d/\d\d\d\d').search(competencia):
                p.alert(text=f'Competência no formato inválido.')
                return
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-input_cnpj_contratante-', True), ('-input_consumer_key-', True), ('-input_consumer_secret-', True), ('-input_senha_certificado-', True), ('-categoria_mensal-', True),
                    ('-categoria_13-', True), ('-declaração-', True), ('-guia_pagamento-', True), ('-input_competencia-', True), ('-abrir-', True), ('-abrir1-', True), ('-abrir2-', True),]:
            window_principal[key[0]].update(disabled=key[1])
            
            # habilita e desabilita os botões conforme necessário
            for key in [('-iniciar-', False), ('-encerrar-', True), ('-abrir_resultados-', True), ('-abrir-', False), ('-abrir1-', False), ('-abrir2-', False),
                        ('-abrir_texto-', True), ('-abrir1_texto-', True), ('-abrir2_texto-', True), ('-progressbar-', True)]:
                window_principal[key[0]].update(visible=key[1])
        
        # controle para saber se os botões estavam visíveis ao trocar o tema da janela
        with open(controle_botoes, 'w', encoding='utf-8') as f:
            f.write('visible')
        
        window['-Mensagens-'].update('Validando credenciais...')
        # atualiza a barra de progresso para ela ficar mais visível
        window['-progressbar-'].update(bar_color=('#fca400', '#ffe0a6'))
        
        try:
            # Chama a função que executa o script
            run(window, cnpj_contratante, usuario_b64, senha, tipo, categoria, competencia, input_certificado, input_excel, output_dir)
        # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            traceback_str = traceback.format_exc()
            
            time.sleep(1)
            if str(erro) == 'Invalid password or PKCS12 data':
                p.alert(text=f'Senha do certificado digital inválida.')
            elif re.compile(r'ConnectTimeoutError').search(str(erro)):
                window['Log do sistema'].update(disabled=False)
                p.alert(text=f'Erro de conexão com o serviço.\n\n'
                             f'Mais detalhes no arquivo "Log.txt" em "Log do sistema"\n')
                escreve_doc(f'Traceback: {traceback_str}\n\n'
                            f'Erro: {erro}', local=output_dir)
            else:
                window['Log do sistema'].update(disabled=False)
                p.alert(text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
                escreve_doc(erro, )
                escreve_doc(f'Traceback: {traceback_str}\n\n'
                            f'Erro: {erro}', local=output_dir)
        
        # habilita e desabilita os botões conforme necessário
        for key in [('-input_cnpj_contratante-', False), ('-input_consumer_key-', False), ('-input_consumer_secret-', False), ('-input_senha_certificado-', False), ('-categoria_mensal-', False),
                    ('-categoria_13-', False), ('-declaração-', False), ('-guia_pagamento-', False), ('-input_competencia-', False), ('-abrir-', False), ('-abrir1-', False), ('-abrir2-', False),
                    ('-iniciar-', False), ('-encerrar-', True)]:
            window_principal[key[0]].update(disabled=key[1])
            
            # habilita e desabilita os botões conforme necessário
            for key in [('-iniciar-', True), ('-encerrar-', False), ('-abrir-', True), ('-abrir1-', True), ('-abrir2-', True),
                        ('-abrir_texto-', False), ('-abrir1_texto-', False), ('-abrir2_texto-', False), ('-progressbar-', False)]:
                window_principal[key[0]].update(visible=key[1])
                
        # apaga qualquer mensagem na interface
        window['-Mensagens-'].update('')
        window['-progressbar-'].update_bar(0)
        window['-Progresso_texto-'].update('')
        """except:
            pass"""
    
    
    window = cria_layout()
    
    categoria = None
    tipo = None
    while True:
        checkboxes_categoria = ['-categoria_mensal-', '-categoria_13-']
        checkboxes_tipo = ['-guia_pagamento-', '-declaração-']
        
        # configura o tema conforme o tema que estava quando o programa foi fechado da última ves
        f = open(dados_modo, 'r', encoding='utf-8')
        modo = f.read()
        if modo == 'claro':
            for key in ['-input_competencia-', '-input_senha_certificado-', '-input_consumer_secret-', '-input_consumer_key-', '-input_cnpj_contratante-']:
                window[key].update(background_color='#D1D1D1')
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-abrir-', '-abrir1-', '-abrir2-']:
                window[key].update(button_color=('#F8F8F8', None))
            window['-tema-'].update(text='☀', button_color=('#FFC100', '#F8F8F8'))
        else:
            for key in ['-input_competencia-', '-input_senha_certificado-', '-input_consumer_secret-', '-input_consumer_key-', '-input_cnpj_contratante-']:
                window[key].update(background_color='#3F3F3F')
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-abrir-', '-abrir1-', '-abrir2-']:
                window[key].update(button_color=('#000000', None))
            window['-tema-'].update(text='🌙', button_color=('#00C9FF', '#000000'))
            
        # captura o evento e os valores armazenados na interface
        event, values = window.read()
        
        # salva o estado da interface
        save_values(values)
        
        # troca o tema
        if event == '-tema-':
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
            
            window.close()  # Fecha a janela atual
            # inicia as variáveis das janelas
            window = cria_layout()
            
            # retorna os elementos preenchidos
            values = load_values()
            for key, value in values.items():
                if key in window.AllKeysDict:
                    try:
                        window[key].update(value)
                    except Exception as e:
                        print(f"Erro ao atualizar o elemento {key}: {e}")
            
            window['-abrir-'].update('Selecione o certificado digital')
            window['-abrir1-'].update('Selecione um arquivo Excel com os dados dos clientes')
            window['-abrir2-'].update('Selecione um diretório para salvar os resultados (Servidor "Comum T" é o padrão):')
            
            # recupera o estado do botão de abrir resultados
            cb = open(controle_botoes, 'r', encoding='utf-8').read()
            if cb == 'visible':
                window['-abrir_resultados-'].update(visible=True)
        
        try:
            cnpj_contratante = values['-input_cnpj_contratante-']
            cnpj_contratante = cnpj_contratante.replace('.', '').replace('/', '').replace('-', '')
            
            consumer_key = values['-input_consumer_key-']
            consumer_secret = values['-input_consumer_secret-']
            
            # concatena os tokens para gerar um único em base64
            usuario = consumer_key + ":" + consumer_secret
            usuario_b64 = converter_base64(usuario)
            senha = values['-input_senha_certificado-']
            competencia = values['-input_competencia-']
            input_certificado = values['-input_certificado-']
            input_excel = values['-input_excel-']
            output_dir = values['-output_dir-']
            contador = 1
            
        except:
            input_excel = 'Desktop'
            output_dir = 'Desktop'
        
        if event in ('-categoria_mensal-', '-categoria_13-'):
            for checkbox in checkboxes_categoria:
                if checkbox != event:
                    window[checkbox].update(value=False)
                else:
                    categoria = checkbox
        
        if event in ('-guia_pagamento-', '-declaração-'):
            for checkbox in checkboxes_tipo:
                if checkbox != event:
                    window[checkbox].update(value=False)
                else:
                    tipo = checkbox
                    
        if event == sg.WIN_CLOSED:
            break
        
        elif event == 'LOG DO SISTEMA':
            os.startfile('Log')

        elif event == 'AJUDA':
            os.startfile('Manual do usuário - Download de documento DCTFWEB API SERPRO.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
        
        elif event == '-abrir_resultados-':
            os.startfile(output_dir)
    
    window.close()
