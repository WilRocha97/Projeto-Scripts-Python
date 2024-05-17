# -*- coding: utf-8 -*-
import traceback, requests_pkcs12, atexit, re, os, requests, time, base64, json, io, chardet, OpenSSL.crypto, contextlib, tempfile, pyautogui as p, PySimpleGUI as sg
from threading import Thread
from pathlib import Path
from functools import wraps

dados = "info\\info.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.readline()
input_consumer = user.split('/')


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
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    # Especifique o caminho para o arquivo de trava
    lock_file_path = 'integra_contador.lock'
    
    # Verifique se outra instância está em execução
    if not create_lock_file(lock_file_path):
        p.alert(text="Outra instância já está em execução.")
        sys.exit(1)
    
    # Defina uma função para remover o arquivo de trava ao final da execução
    atexit.register(remove_lock_file, lock_file_path)
    
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    categoria_key = ('GERAL_MENSAL', 'GERAL_13o_SALARIO')
    categoria_nome = ('Mensal', '13º')
    # sg.theme_previewer()
    # Layout da janela
    layout = [
        [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0)],
        [sg.Text('')],
        [sg.Text('Informe o CNPJ do contratante do serviço SERPRO:')],
        [sg.InputText(key='-input_cnpj_contratante-', size=90, default_text=input_consumer[0])],
        [sg.Text('Informe a Consumer Key:')],
        [sg.InputText(key='-input_consumer_key-', size=90, password_char='*', default_text=input_consumer[1])],
        [sg.Text('Informe a Consumer Secret:')],
        [sg.InputText(key='-input_consumer_secret-', size=90, password_char='*', default_text=input_consumer[2])],
        [sg.Text('Informe a senha do certificado digital:')],
        [sg.InputText(key='-input_senha_certificado-', size=90, password_char='*')],
        [sg.Checkbox(key='-categoria_mensal-', text='Mensal', enable_events=True), sg.Checkbox(key='-categoria_13-', text='13º', enable_events=True)],
        [sg.Checkbox(key='-guia_pagamento-', text='Documento de arrecadação', enable_events=True), sg.Checkbox(key='-declaração-', text='Declaração Completa', enable_events=True)],
        [sg.Text('Informe a competência das guias. (Mensal = "00/0000") (13º = "0000"):')],
        [sg.InputText(key='-input_competencia-', size=90)],
        [sg.Text('')],
        [sg.Text('Selecione o certificado digital:')],
        [sg.FileBrowse('Pesquisar', key='-abrir-', file_types=(('PFX files', '*.pfx'),)), sg.InputText(key='-input_certificado-', size=80, disabled=True)],
        [sg.Text('Selecione um arquivo Excel com os dados dos clientes:')],
        [sg.FileBrowse('Pesquisar', key='-abrir1-', file_types=(('Planilhas Excel', '*.csv'),)), sg.InputText(key='-input_excel-', size=80, disabled=True)],
        [sg.Text('Selecione um diretório para salvar os resultados (Servidor "Comum T:" é o padrão):')],
        [sg.FolderBrowse('Pesquisar', key='-abrir2-'), sg.InputText(default_text='T:/ROBÔ/DCTF-WEB',key='-output_dir-', size=80, disabled=True)],
        [sg.Text('')],
        [sg.Text('', key='-Mensagens-')],
        [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color='#f0f0f0')],
        [sg.Button('Iniciar', key='-iniciar-', border_width=0), sg.Button('Encerrar', key='-encerrar-', disabled=True, border_width=0), sg.Button('Abrir resultados', key='-abrir_resultados-', disabled=True, border_width=0)],
    ]
    
    # guarda a janela na variável para manipula-la
    window = sg.Window('Download de documentos DCTFWEB API SERPRO', layout)
    
    
    def run_script_thread():
        # try:
        if not cnpj_contratante:
            p.alert(text=f'Por favor informe o CNPJ do contratante da API SERPRO.')
            return
        if not len(cnpj_contratante) == 14:
            p.alert(text=f'Por favor informe um CNPJ válido.')
            return
        if not consumer_key:
            p.alert(text=f'Por favor informe o consumerKey.')
            return
        if not consumer_secret:
            p.alert(text=f'Por favor informe o consumerSecret.')
            return
        if not senha:
            p.alert(text=f'Por favor informe a senha do certificado digital.')
            return
        if not categoria:
            p.alert(text=f'Por favor informe se a categoria é Mensal ou 13º.')
            return
        if not tipo:
            p.alert(text=f'Por favor informe se o tipo de documento é o Documento de Arrecadação ou Declaração completa')
            return
        if not competencia:
            p.alert(text=f'Por favor informe a competência das guias.')
            return
        if not input_certificado:
            p.alert(text=f'Por favor selecione um certificado digital.')
            return
        if not input_excel:
            p.alert(text=f'Por favor selecione uma planilha de dados.')
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
        window['-input_cnpj_contratante-'].update(disabled=True)
        window['-input_consumer_key-'].update(disabled=True)
        window['-input_consumer_secret-'].update(disabled=True)
        window['-input_senha_certificado-'].update(disabled=True)
        window['-categoria_mensal-'].update(disabled=True)
        window['-categoria_13-'].update(disabled=True)
        window['-declaração-'].update(disabled=True)
        window['-guia_pagamento-'].update(disabled=True)
        window['-input_competencia-'].update(disabled=True)
        window['-abrir-'].update(disabled=True)
        window['-abrir1-'].update(disabled=True)
        window['-abrir2-'].update(disabled=True)
        window['-iniciar-'].update(disabled=True)
        window['-encerrar-'].update(disabled=False)
        window['-abrir_resultados-'].update(disabled=False)
        
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
        window['-input_cnpj_contratante-'].update(disabled=False)
        window['-input_consumer_key-'].update(disabled=False)
        window['-input_consumer_secret-'].update(disabled=False)
        window['-input_senha_certificado-'].update(disabled=False)
        window['-categoria_mensal-'].update(disabled=False)
        window['-categoria_13-'].update(disabled=False)
        window['-declaração-'].update(disabled=False)
        window['-guia_pagamento-'].update(disabled=False)
        window['-input_competencia-'].update(disabled=False)
        window['-abrir-'].update(disabled=False)
        window['-abrir1-'].update(disabled=False)
        window['-abrir2-'].update(disabled=False)
        window['-iniciar-'].update(disabled=False)
        window['-encerrar-'].update(disabled=True)
        
        # apaga qualquer mensagem na interface
        window['-Mensagens-'].update('')
        # atualiza a barra de progresso para ela ficar mais visível
        window['-progressbar-'].update_bar(0)
        window['-Progresso_texto-'].update('')
        window['-progressbar-'].update(bar_color='#f0f0f0')
        """except:
            pass"""
    
    
    categoria = None
    tipo = None
    while True:
        checkboxes_categoria = ['-categoria_mensal-', '-categoria_13-']
        checkboxes_tipo = ['-guia_pagamento-', '-declaração-']
        # captura o evento e os valores armazenados na interface
        event, values = window.read()
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
        
        elif event == 'Log do sistema':
            os.startfile('Log')

        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Download de documento DCTFWEB API SERPRO.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
        
        elif event == '-abrir_resultados-':
            os.startfile(output_dir)
    
    window.close()