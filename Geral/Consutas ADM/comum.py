# -*- coding: utf-8 -*-
import chromedriver_autoinstaller, time, shutil, re, os, pandas as pd
from datetime import timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from anticaptchaofficial.imagecaptcha import *
from anticaptchaofficial.recaptchav2proxyless import *
from anticaptchaofficial.hcaptchaproxyless import *
from pyautogui import alert
from PIL import Image

dados_anticaptcha = "Dados/Dados AC.txt"
dados_contadores = "Dados/Contadores.txt"
controle_rotinas = 'Log/Controle.txt'


def concatena(variavel, quantidade, posicao, caractere):
    # função para concatenar strings colocando caracteres no começo ou no final
    variavel = str(variavel)
    if posicao == 'depois':
        # concatena depois
        while len(str(variavel)) < quantidade:
            variavel += str(caractere)
    if posicao == 'antes':
        # concatena antes
        while len(str(variavel)) < quantidade:
            variavel = str(caractere) + str(variavel)
    
    return variavel


def configura_navegador(window_principal, pasta=None, retorna_options=False, timeout=90):
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    # options.add_argument("--start-maximized")
    if pasta:
        options.add_experimental_option('prefs', {
            "download.default_directory": pasta.replace('/', '\\'),  # Change default directory for downloads
            "download.prompt_for_download": False,  # To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,  # It will not show PDF directly in chrome
            "profile.default_content_setting_values.automatic_downloads": 1  # download multiple files
        })
    
    if retorna_options:
        return options
    
    window_principal['-Mensagens-'].update('Iniciando ambiente da consulta, aguarde...')
    window_principal.refresh()
    return initialize_chrome(timeout, options)


def initialize_chrome(timeout, options=webdriver.ChromeOptions()):
    service = None
    shutil.rmtree('Chrome driver')
    time.sleep(1)
    os.makedirs('Chrome driver', exist_ok=True)
    
    # biblioteca para baixar o chromedriver atualizado
    chromedriver_autoinstaller.install(path='Chrome driver')
    print('>>> Inicializando Chromedriver...')
    
    for pasta_atual, subpastas, arquivos in os.walk('Chrome driver'):
        # Agora você pode processar os arquivos na pasta atual normalmente
        for file in arquivos:
            caminho_completo = os.path.join(pasta_atual, file)
            service = Service(caminho_completo)
            print(caminho_completo)
            
    if not service:
        return False, 'Não encontrou o chrome driver'
    
    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
    
    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # retorna o chromedriver aberto
    driver = webdriver.Chrome(options=options, service=service)
    driver.set_page_load_timeout(timeout)
    return True, driver


def find_by_id(item, driver):
    try:
        driver.find_element(by=By.ID, value=item)
        return True
    except:
        return False


def find_by_path(item, driver):
    try:
        driver.find_element(by=By.XPATH, value=item)
        return True
    except:
        return False


def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass


def send_input_xpath(elem_path, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_path)
            elem.send_keys(data)
            break
        except:
            pass


def get_info_post(soup):
    # captura infos para realizar um post
    soup = BeautifulSoup(soup, 'html.parser')
    infos = [
        soup.find('input', attrs={'id': '__VIEWSTATE'}),
        soup.find('input', attrs={'id': '__VIEWSTATEGENERATOR'}),
        soup.find('input', attrs={'id': '__EVENTVALIDATION'})
    ]
    
    # state, generator, validation
    return tuple(info.get('value', '') for info in infos if info)


def new_session_fazenda_driver(window_principal, user, pwd, perfil, retorna_driver=False, options='padrão'):
    url_home = "https://www3.fazenda.sp.gov.br/CAWEB/Account/Login.aspx"
    _site_key = '6LesbbcZAAAAADrEtLsDUDO512eIVMtXNU_mVmUr'
    
    window_principal['-Mensagens_2-'].update('Iniciando ambiente da consulta, aguarde...')
    window_principal.refresh()
    
    if options == 'padrão':
        # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--window-size=1920,1080')
        # options.add_argument("--start-maximized")
    
    status, driver = initialize_chrome(options)
    driver.get(url_home)
    
    # gera o token para passar pelo captcha
    recaptcha_data = {'sitekey': _site_key, 'url': url_home}
    token = solve_recaptcha(recaptcha_data)
    if not token:
        return False, 'erro_captcha'
    token = str(token)
    
    
    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
    if cr == 'STOP':
        return False, ''
    
    if perfil == 'contador':
        window_principal['-Mensagens_2-'].update(f'Abrindo perfil do contador...')
        window_principal.refresh()
        button = driver.find_element(by=By.ID, value='ConteudoPagina_rdoListPerfil_1')
        button.click()
        time.sleep(1)
    
    elif perfil == 'contribuinte':
        window_principal['-Mensagens_2-'].update(f'Abrindo perfil do contribuinte...')
        window_principal.refresh()
    
    print(f'>>> Logando no usuário')
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtUsuario')
    element.send_keys(user)
    
    element = driver.find_element(by=By.ID, value='ConteudoPagina_txtSenha')
    element.send_keys(pwd)
    
    script = 'document.getElementById("g-recaptcha-response").innerHTML="{}";'.format(token)
    driver.execute_script(script)
    
    script = 'document.getElementById("ConteudoPagina_btnAcessar").disabled = false;'
    driver.execute_script(script)
    time.sleep(1)
    
    button = driver.find_element(by=By.ID, value='ConteudoPagina_btnAcessar')
    button.click()
    time.sleep(3)
    
    try:
        button = driver.find_element(by=By.XPATH, value='/html/body/div[2]/section/div/div/div/div[2]/div/ul/li/form/div[5]/div/a')
        button.click()
        time.sleep(3)
    except:
        pass
    
    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    padrao = re.compile(r'SID=(.\d+)')
    resposta = padrao.search(str(soup))
    
    if not resposta:
        try:
            padrao = re.compile(r'(Senha inserida está incorreta)')
            resposta = padrao.search(str(soup))
            
            if not resposta:
                padrao = re.compile(r'(Favor informar o login e a senha corretamente.)')
                resposta = padrao.search(str(soup))
                
                if not resposta:
                    padrao = re.compile(r'(O usuário não tem acesso a este serviço do sistema ou o serviço não foi liberado para a sua utilização.)')
                    resposta = padrao.search(str(soup))
                    
                    if not resposta:
                        padrao = re.compile(r'(ERRO INTERNO AO SISTEMA DE CONTROLE DE ACESSO)')
                        driver.save_screenshot(r'ignore\debug_screen.png')
                        resposta = padrao.search(str(soup))
            
            sid = resposta.group(1)
            print(f'❌ {sid}')
            cokkies = 'erro'
            driver.close()
            
            return cokkies, sid
        except:
            driver.save_screenshot(r'Log\debug_screen.png')
            driver.close()
            return False, 'Erro ao logar no perfil'
    
    if retorna_driver:
        sid = resposta.group(1)
        return driver, sid
    
    sid = resposta.group(1)
    cookies = driver.get_cookies()
    driver.quit()
    
    return cookies, sid


def solve_recaptcha(data):
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_anti = f.read()
    anticaptcha_api_key = key_anti
    
    print('>>> Quebrando recaptcha')
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    solver.set_website_url(data['url'])
    solver.set_website_key(data['sitekey'])
    # set optional custom parameter which Google made for their search page Recaptcha v2
    # solver.set_data_s('"data-s" token from Google Search results "protection"')
    
    g_response = solver.solve_and_return_solution()
    if g_response != 0:
        return g_response
    
    elif solver.error_code == 'ERROR_ZERO_BALANCE':
        alert('❌ Não consta saldo para resolução do captcha, mande um e-mail solicitando a recarga para o desenvolvedor.')
        print("Erro: Saldo insuficiente na conta AntiCaptcha.")
        return False
        
    else:
        print(solver.error_code)
        return solver.error_code


def solve_text_captcha(driver, captcha_element, element_type='id'):
    element = ''
    
    f = open(dados_anticaptcha, 'r', encoding='utf-8')
    key_anti = f.read()
    anticaptcha_api_key = key_anti
    
    os.makedirs('Log\captcha', exist_ok=True)
    # captura a imagem do captcha
    if element_type == 'id':
        element = driver.find_element(by=By.ID, value=captcha_element)
    elif element_type == 'xpath':
        element = driver.find_element(by=By.XPATH, value=captcha_element)
    
    location = element.location
    size = element.size
    driver.save_screenshot('Log\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    time.sleep(2)
    im = Image.open(r'Log\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'Log\captcha\captcha.png')
    time.sleep(1)
    
    print('>>> Quebrando text captcha')
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    
    captcha_text = solver.solve_and_return_solution(os.path.join('Log', 'captcha', 'captcha.png'))
    if captcha_text != 0:
        return captcha_text
    
    elif solver.error_code == 'ERROR_ZERO_BALANCE':
        alert('❌ Não consta saldo para resolução do captcha, mande um e-mail solicitando a recarga para o desenvolvedor.')
        print("Erro: Saldo insuficiente na conta AntiCaptcha.")
        return False
    
    else:
        print(solver.error_code)
        return solver.error_code


def indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao):
    # print(count, 'primeiro')
    tempo_estimado = 0
    if type(total_empresas) == list:
        quantidade_total_empresas = len(total_empresas)
    else:
        quantidade_total_empresas = int(total_empresas)
    
    # captura a hora atual e coloca em uma lista para calcular o tempo de execução do andamento atual
    tempos.append(datetime.now())
    tempo_execucao_atual = int(tempos[1].timestamp()) - int(tempos[0].timestamp())
    
    # adiciona o tempo de execução atual na lista com os tempos anteriores para calcular a média de tempo de execução dos andamentos
    tempo_execucao.append(tempo_execucao_atual)
    for t in tempo_execucao:
        tempo_estimado = tempo_estimado + t
    tempo_estimado = int(tempo_estimado) / int(len(tempo_execucao))
    
    # multiplica o tempo médio de execução dos andamentos pelo número de andamentos que faltam executar para obter o tempo estimado em segundos
    tempo_total_segundos = int((quantidade_total_empresas + index) - (count + index) + 1) * float(tempo_estimado)
    # Converter o tempo total para um objeto timedelta
    tempo_total = timedelta(seconds=tempo_total_segundos)
    
    # Extrair dias, horas e minutos do timedelta
    dias = tempo_total.days
    horas = tempo_total.seconds // 3600
    minutos = (tempo_total.seconds % 3600) // 60
    segundos = (tempo_total.seconds % 3600) % 60
    
    # Retorna o tempo no formato "dias:horas:minutos:segundos"
    dias_texto = ''
    horas_texto = ''
    minutos_texto = ''
    segundos_texto = ''
    
    if dias > 0:
        dias_texto = f' {dias} dias'
    if horas > 0:
        horas_texto = f' {horas} horas'
    if minutos > 0:
        minutos_texto = f' {minutos} minutos'
    if segundos > 0:
        segundos_texto = f' {segundos} segundos'
    
    tempo_estimado_texto = f" | Tempo estimado:{dias_texto}{horas_texto}{minutos_texto}{segundos_texto}"
    
    tempos.pop(0)
    
    print(f'\n\n[{empresa}]')
    # print(count, index, 'segundo')
    window_principal['-progressbar-'].update(visible=True)
    window_principal['-Mensagens-'].update(f'{str((count + index) - 1)} de {str(quantidade_total_empresas)} | {str(quantidade_total_empresas - (count + index) + 1)} Restantes{tempo_estimado_texto}')
    window_principal['-progressbar-'].update_bar(count - 1, max=int(quantidade_total_empresas) - int(index))
    window_principal['-Progresso_texto-'].update(str(round(float(count - 1) / (int(quantidade_total_empresas) - int(index)) * 100, 1)) + '%')
    window_principal.refresh()
    print(f'{str((count + index) - 1)} de {str(quantidade_total_empresas)} | {str(quantidade_total_empresas - (count + index) + 1)} Restantes{tempo_estimado_texto}')
    
    tempo_estimado = tempo_execucao
    return tempos, tempo_estimado


def open_dados(situacao_dados, andamentos, empresas_20000, pasta_final, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, filtrar_celulas_em_branco):
    dados_final = os.path.join(pasta_final, 'Dados.xlsx')
    encode = 'latin-1'
    
    # modelo de lista com as colunas que serão usadas na rotina
    # colunas_usadas = ['column1', 'column2', 'column3']
    
    if situacao_dados == '-nova_planilha-':
        df = pd.read_excel(planilha_dados)
        
        # coluna com os códigos do ae
        coluna_codigo = 'Codigo'
        
        if empresas_20000 == 'Empresas com o código menor que 20.000':
            # cria um novo df apenas com empresas a baixo do código 20.000
            df_filtrada = df[df[coluna_codigo] <= 20000]
        elif empresas_20000 == 'Empresas com o código maior que 20.000':
            # cria um novo df apenas com empresas a cima do código 20.000
            df_filtrada = df[df[coluna_codigo] >= 20000]
        else:
            df_filtrada = df
        
        # filtra as células de colunas específicas que contenham palavras especificas
        if palavras_filtro and colunas_filtro:
            for count, coluna_para_filtrar in enumerate(colunas_filtro):
                df_filtrada = df_filtrada[df_filtrada[coluna_para_filtrar].str.contains(palavras_filtro[count], case=False, na=False)]
        
        # filtra as colunas
        try:
            df_filtrada = df_filtrada[colunas_usadas]
        except KeyError:
            alert(f'❌ Erro ao buscar as colunas na planilha base selecionada: {planilha_dados}\n\n'
                  f'Verifique se a planilha contem as colunas necessárias para a execução da rotina e se elas tem exatamente o mesmo nome indicado ao lado: {colunas_usadas}')
            return False
        
        if filtrar_celulas_em_branco:
            df_filtrada = df_filtrada.dropna(subset=filtrar_celulas_em_branco)
            # df_filtrada = df_filtrada.fillna('vazio')
        else:
            # remove linha com células vazias
            df_filtrada = df_filtrada.dropna(axis=0, how='any')
        
        # Converte a coluna 'CNPJ' para string e remova a parte decimal '.0'. Preencha com zeros à esquerda para garantir 14 dígitos
        df_filtrada['CNPJ'] = df_filtrada['CNPJ'].astype(str).str.replace(r'\.0', '', regex=True).str.zfill(14)
        
        if andamentos == 'Consulta Débitos Estaduais - Situação do Contribuinte' or andamentos == 'Consulta Certidão Negativa de Débitos Tributários Não Inscritos':
            contadores_dict = atualiza_contadores()
            # Substituir valores com base no dicionário apenas se o valor estiver presente no dicionário
            df_filtrada['Perfil'] = 'vazio'
            
            # Função para atualizar os valores das colunas com base no dicionário de mapeamento
            def atualizar_valores(row):
                if row['PostoFiscalContador'] in contadores_dict:
                    return contadores_dict[row['PostoFiscalContador']]
                else:
                    return (row['PostoFiscalUsuario'], row['PostoFiscalSenha'], 'contribuinte')
            
            # Aplicar a função para atualizar os valores das colunas
            df_filtrada[['PostoFiscalUsuario', 'PostoFiscalSenha', 'Perfil']] = df_filtrada.apply(atualizar_valores, axis=1, result_type='expand')
            
            # 5. Deletar a coluna 'contador'
            df_filtrada.drop(columns=['PostoFiscalContador'], inplace=True)
            
            # 3. Deletar linhas com células vazias na coluna 'senha'
            df_filtrada = df_filtrada.dropna(subset=['PostoFiscalSenha'])
            
            # Ordene o DataFrame com base na coluna desejada
            df_filtrada = df_filtrada.sort_values(by=['Perfil', 'PostoFiscalUsuario', 'CNPJ'], ascending=[True, True, True])
            
            # remove linha com células vazias
            df_filtrada = df_filtrada.dropna(axis=0, how='any')
            
            # Remover linhas que contenham 'ISENTO' na coluna 'PostoFiscalUsuario'
            df_filtrada = df_filtrada[~df_filtrada['PostoFiscalUsuario'].str.contains('ISENTO', case=False, na=False)]
            # Remover linhas que contenham 'BAIXADO' na coluna 'PostoFiscalUsuario'
            df_filtrada = df_filtrada[~df_filtrada['PostoFiscalUsuario'].str.contains('BAIXADO', case=False, na=False)]
        
        if df_filtrada.empty:
            alert(f'❗ Não foi encontrado nenhuma empresa na planilha selecionada: {planilha_dados}\n\n'
                  f'{empresas_20000}\n'
                  f'utilizando os seguintes filtros: {palavras_filtro}\n'
                  f'nas respectivas colunas {colunas_filtro}\n')
            return False
        
        for coluna in df_filtrada.columns:
            # Remova aspas duplas
            df_filtrada[coluna] = df_filtrada[coluna].str.replace('"', '')
            
            # Remova quebras de linha (`\n` e `\r`)
            df_filtrada[coluna] = df_filtrada[coluna].str.replace('\n', '').str.replace('\r', '').str.replace('_x000D_', '')
        
        df_filtrada.to_excel(dados_final, index=False)
        empresas = pd.read_excel(dados_final)
    else:
        empresas = pd.read_excel(planilha_dados)
    
    print(empresas)
    return empresas


def escreve_doc(texto, local='Log', nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    try:
        os.remove(os.path.join(local, 'Log.txt'))
    except:
        pass
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()


def escreve_relatorio_xlsx(texto, local, nome='Relatório', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        df_status = pd.read_excel(os.path.join(local, f"{nome}.xlsx"))
        df_status.loc[len(df_status)] = texto
        with pd.ExcelWriter(os.path.join(local, f"{nome}.xlsx"), engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_status.to_excel(writer, index=False, sheet_name='Sheet1')  # Certifique-se de que o nome da planilha esteja correto
    except:
        df_status = pd.DataFrame(texto, index=[0])
        df_status.to_excel(os.path.join(local, f"{nome}.xlsx"), index=False)


def escreve_header(planilha_resultado, header):
    # abre a planilha de andamentos
    df = pd.read_csv(planilha_resultado)
    # definir o novo cabeçalho
    df.columns = header
    # salva a nova planilha
    df.to_excel(planilha_resultado.replace('.csv', '.xlsx'), index=False)
    

def configura_dados(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, colunas_usadas=None, colunas_filtro=None, palavras_filtro=None, filtrar_celulas_em_branco=None):
    def where_to_start(pasta_final_anterior, planilha_andamentos, df_empresas):
        if not os.path.isdir(pasta_final_anterior):
            return 0
        
        try:
            df_andamentos = pd.read_excel(os.path.join(pasta_final_anterior, planilha_andamentos))
        except:
            alert(f'❗ Não foi encontrada nenhuma planilha de andamentos na pasta de execução anterior.\n\n'
                  f'Começando a execução a partir do primeiro índice da planilha de dados selecionada.')
            return 0
        
        # pega o valor da última linha da primeira coluna para buscar o index na planilha de dados
        ultima_linha_processada = df_andamentos.iloc[-1, 0]
        
        # Procurar esse valor na primeira coluna do segundo DataFrame
        index = df_empresas[df_empresas.iloc[:, 0] == ultima_linha_processada].index
        
        # Se última linha processada não for encontrada, iniciar do começo
        if not index.empty:
            return int(index[0]) + 1
        else:
            return 0
    
    comp = datetime.now().strftime('%m-%Y')
    pasta_final_ = os.path.join(pasta_final_, andamentos, comp)
    contador = 0
    if planilha_dados == 'Não se aplica':
        empresas_20000 = ''
    else:
        if codigo_20000 == '-codigo_20000_nao-':
            empresas_20000 = ' - (Empresas com o código menor que 20.000)'
        elif codigo_20000 == '-codigo_20000-':
            empresas_20000 = ' - (Empresas com o código maior que 20.000)'
        else:
            empresas_20000 = ' - (Empresas com qualquer código)'
       
    # iteração para determinar se precisa criar uma pasta nova para armazenar os resultados
    # toda vês que o programa começar as consultas uma nova pasta será criada para não sobrepor ou misturar as execuções
    while True:
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'STOP':
            return '', '', False
        try:
            os.makedirs(os.path.join(pasta_final_, f'Execução{empresas_20000}'))
            pasta_final = os.path.join(pasta_final_, f'Execução{empresas_20000}')
            pasta_final_anterior = False
            break
        except:
            try:
                contador += 1
                os.makedirs(os.path.join(pasta_final_, f'Execução{empresas_20000} ({str(contador)})'))
                pasta_final = os.path.join(pasta_final_, f'Execução{empresas_20000} ({str(contador)})')
                if contador - 1 < 1:
                    pasta = f'Execução{empresas_20000}'
                else:
                    pasta = f'Execução{empresas_20000} ({str(contador - 1)})'
                pasta_final_anterior = os.path.join(pasta_final_, pasta)
                break
            except:
                pass

    if planilha_dados != 'Não se aplica':
        # abrir a planilha de dados
        window_principal['-Mensagens_2-'].update('Criando dados para a consulta...')
        window_principal.refresh()
        df_empresas = open_dados(situacao_dados, andamentos, empresas_20000, pasta_final, planilha_dados, colunas_usadas, colunas_filtro, palavras_filtro, filtrar_celulas_em_branco)
        if df_empresas.empty:
            return '', '', False
        
        if situacao_dados == '-minha_planilha-':
            index = 0
        elif pasta_final_anterior:
            planilha_andamentos = f'{andamentos}.xlsx'
            # obtêm o índice do último andamento da execução anterior para continuar
            index = where_to_start(pasta_final_anterior, planilha_andamentos, df_empresas)
            print(index)
        else:
            index = 0
            
        total_empresas = int(df_empresas.shape[0])
    else:
        index = None
        df_empresas = None
        total_empresas = None
    
    return pasta_final, index, df_empresas, total_empresas


def atualiza_contadores():
    # obtêm a lista de usuário e senha de cada contador para a planilha dedados de algumas consultas
    f = open(dados_contadores, 'r', encoding='utf-8')
    contadores = f.readlines()
    
    contadores_dict = {}
    for contador in contadores:
        contador = contador.split('/')
        contadores_dict[contador[0]] = (contador[1], contador[2], contador[3])
    
    return contadores_dict


def download_file(name, response, pasta):
    # função para salvar um arquivo retornado de uma requisição
    pasta = str(pasta).replace('\\', '/')
    os.makedirs(pasta, exist_ok=True)
    
    with open(os.path.join(pasta, name), 'wb') as arq:
        for i in response.iter_content(100000):
            arq.write(i)
