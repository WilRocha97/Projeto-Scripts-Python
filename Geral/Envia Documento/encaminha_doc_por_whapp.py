# -*- coding: utf-8 -*-
import datetime, os, random, time, re, pandas as pd, pyautogui as p

import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from pathlib import Path

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers, _remove_emojis, _barra_de_status
from chrome_comum import _abrir_chrome, _initialize_chrome, _find_by_id
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img


e_dir = Path('//vpsrv03/Arq_Robo/Envia WP por e-mail/Execução')
dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados e-mail.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')
email = user[0]
senha = user[1]


def login_email(driver):
    driver.get('https://smartermail.hiway.com.br/')
    print('>>> Acessando o site')
    time.sleep(0.5)
    while True:
        try:
            driver.find_element(by=By.ID, value='loginUsernameBox').send_keys(email)
            driver.find_element(by=By.ID, value='loginPasswordBox').send_keys(senha)
            break
        except:
            pass
        
    time.sleep(0.5)
    
    driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/div[1]/div[1]/form/div[2]/div[2]/button').click()
    
    _wait_img('menu_classificacao.png', conf=0.9)
    _click_img('menu_classificacao.png', conf=0.9)
    _wait_img('crescente.png', conf=0.9)
    time.sleep(0.5)
    _click_img('crescente.png', conf=0.9)
    time.sleep(1)
    
    return driver


def captura_dados_email(driver):
    print('>>> Capturando dados da mensagem')
    
    while not _find_by_id('messageViewFrame', driver):
        print('aqui')
        time.sleep(1)

    while True:
        try:
            titulo = re.compile(r'subject\">(.+)</div>').search(driver.page_source).group(1)
            titulo = titulo.replace('-&nbsp;', '- ').replace(' &nbsp;', ' ').replace('&nbsp; ', ' ').replace('&nbsp;', ' ').replace('&amp;', '&')
            break
        except:
            print('aqui2')

    print(titulo)
    pasta_anexos = 'V:\\Setor Robô\\Scripts Python\\Geral\\Envia Documento\\ignore\\Anexos'
    
    if titulo[:3] == 'Re:':
        return driver, titulo, 'x', 'x', 'x', 'x', 'x'
    
    if _find_img('anexos.png', conf=0.9):
        for arq in os.listdir(pasta_anexos):
            os.remove(os.path.join(pasta_anexos, arq))

        driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/div[2]/div/div/div/div/div[2]/div/div/div[2]/div/div[3]/div[3]/message-header/div[3]/div/div/a').click()
        _wait_img('baixar_anexo.png', conf=0.9)
        _click_position_img('baixar_anexo.png', '+', pixels_y=50, conf=0.9)
        
        while len(os.listdir(pasta_anexos)) < 1:
            time.sleep(1)
        time.sleep(2)
        
        for arq in os.listdir(pasta_anexos):
            anexo = os.path.join('V:\\', 'Setor Robô', 'Scripts Python', 'Geral', 'Envia Documento', 'ignore', 'Anexos', arq)
        
        _click_img('fechar_anexos.png', conf=0.9)
        
        # Encontra o elemento do frame pelo nome ou índice
        frame = driver.find_element(by=By.ID, value='messageViewFrame')
        # frame = driver.find_element_by_index(0)
        
        # Alterna o driver para o contexto do frame
        driver.switch_to.frame(frame)
        # print(driver.page_source)
        mensagem_email = re.compile(r'\[(.+)].+Prezado cliente, tenha um excelente dia.<br> <br>(.+),<br><br><\/div>').search(driver.page_source)
        cnpj = mensagem_email.group(1)
        cnpj_limpo = cnpj
        
        corpo_email = mensagem_email.group(2)
        corpo_email = corpo_email.replace('<br>', '\n').replace('desse e-mail', 'dessa conversa').replace('este e-mail', 'esta conversa')
        
        return driver, titulo, cnpj, cnpj_limpo, corpo_email, '', anexo
        
    # Encontra o elemento do frame pelo nome ou índice
    frame = driver.find_element(by=By.ID, value='messageViewFrame')
    # frame = driver.find_element_by_index(0)
    
    # Alterna o driver para o contexto do frame
    driver.switch_to.frame(frame)
    # print(driver.page_source)
    try:
        # pega cnpj da empresa que vai receber a mensagem
        cnpj = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)').search(driver.page_source).group(1)
        cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
        print(cnpj)
        print(cnpj_limpo)
    except:
        try:
            # pega cnpj da empresa que vai receber a mensagem
            cnpj = re.compile(r'CNPJ:</strong> (\d\d\d\d\d\d\d\d\d\d\d)&nbsp;').search(driver.page_source).group(1)
            cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
            print(cnpj)
            print(cnpj_limpo)
        except:
            try:
                # pega cnpj da empresa que vai receber a mensagem
                cnpj = re.compile(r'CNPJ:</strong> (\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)&nbsp;').search(driver.page_source).group(1)
                cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
                print(cnpj)
                print(cnpj_limpo)
            except:
                try:
                    # pega cnpj da empresa que vai receber a mensagem
                    cnpj = re.compile(r'CNPJ: (\d\d\d\d\d\d\d\d\d\d\d)&nbsp').search(driver.page_source).group(1)
                    cnpj_limpo = cnpj.replace('.', '').replace('/', '').replace('-', '')
                    print(cnpj)
                    print(cnpj_limpo)
                except:
                    return driver, titulo, 'x', 'x', 'x', 'x', 'x'
        
    # pega a data de vencimento do documento que está no link da mensagem
    try:
        vencimento = re.compile(r'Vencimento:.+(\d\d/\d\d/\d\d\d\d)').search(driver.page_source).group(1)
        vencimento = f"Vencimento: {vencimento}"
    except:
        vencimento = ''
    print(vencimento)
    
    # pega o link que será enviado na mensagem
    link_email_padrao = re.compile(r'<a href=\"(https://api\.gestta\.com\.br.+)\" style=\"margin: 0; padding: 0; font-family:').findall(driver.page_source)
    if not link_email_padrao:
        link_email_padrao = re.compile(r'<a href=\"(https://app\.gestta\.com\.br.+)\" style=\"margin: 0; padding: 0; font-family:').findall(driver.page_source)
        if not link_email_padrao:
            link_email_padrao = re.compile(r'href=\"(https://forms.office.com.+)\">https:').findall(driver.page_source)
    
    link_email = list(set(link_email_padrao))
    
    link_mensagem = ''
    contador = 1
    for link in link_email:
        link_mensagem += str(contador) + ' - ' + link + '\n'
        contador += 1
    
    link_mensagem = link_mensagem.replace(' ', '').replace('&amp;', '&')
    print(link_mensagem)
    return driver, titulo, cnpj, cnpj_limpo, None, vencimento, str(link_mensagem)


def verifica_o_numero(cnpj_pesquisado):
    try:
        print('>>> Consultando número de telefone do cliente')
        # Definir o nome das colunas
        colunas = ['cnpj', 'numero', 'nome']
        # Localiza a planilha
        caminho = Path('T:/ROBO/Envia WP por e-mail/Dados/Dados.csv')
        # Abrir a planilha
        lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
        # Definir o index da planilha
        lista.set_index('cnpj', inplace=True)
        cliente = lista.loc[int(cnpj_pesquisado)]
        
        try:
            cnpjs_iguais = []
            for key in cliente.itertuples():
                cnpjs_iguais.append({'cnpj': key[0], 'numero': key[1], 'nome': key[2]})
        except:
            cliente = list(cliente)
            numero = cliente[0]
            nome = cliente[1]
            cnpjs_iguais = [{'cnpj': cnpj_pesquisado, 'numero': numero, 'nome': nome}]
            
        print(cnpjs_iguais)
        return cnpjs_iguais
    except:
        return False


def envia(resultado_anterior, contato, titulo, vencimento, link_mensagem, corpo_email, nome_planilha):
    cnpj = str(contato['cnpj'])
    nome = str(contato['nome'])
    numero = str(contato['numero'])
    titulo_sem_emoji = _remove_emojis(titulo)
    print('>>> Enviando mensagem')
    if numero == 'numero':
        return 'erro'
    try:
        if not corpo_email:
            resultado = enviar_sem_anexo(numero, link_mensagem, titulo, vencimento)
            """pywhatkit.sendwhatmsg_instantly('+55' + numero, f"Olá!\n"
                                                            f"{titulo}\n"
                                                            f"{vencimento}\n"
                                                            f"{link_mensagem}\n"
                                                            f"Obrigado,\n"
                                                            f"R.POSTAL SERVICOS CONTABEIS LTDA\n"
                                                            f"veigaepostal@veigaepostal.com.br\n"
                                                            f"(19)3829-8959", 50, True, 10)"""
        else:
            resultado = enviar_anexo(numero, link_mensagem, corpo_email)
            
        time.sleep(1)
        if resultado != 'ok':
            try:
                _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo};{resultado}', nome=nome_planilha + ' erros', local=e_dir)
            except:
                _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo_sem_emoji};{resultado}', nome=nome_planilha + ' erros', local=e_dir)
            print(f'❌ {resultado}\n')
            if resultado_anterior == 'ok':
                return 'ok'
            else:
                return resultado
        try:
            _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo};Mensagem enviada', nome=nome_planilha, local=e_dir)
        except:
            _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo_sem_emoji};Mensagem enviada', nome=nome_planilha, local=e_dir)
            
        print('✔ Mensagem enviada\n')
        return 'ok'

    except:
        try:
            _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo};Erro ao enviar a mensagem', nome=nome_planilha + ' erros', local=e_dir)
        except:
            _escreve_relatorio_csv(f'{cnpj};{nome};{numero};{titulo_sem_emoji};Erro ao enviar a mensagem', nome=nome_planilha + ' erros', local=e_dir)
        print('❌ Erro ao enviar a mensagem\n')
        if resultado_anterior == 'ok':
            return 'ok'
        else:
            return 'erro'


def enviar_sem_anexo(numero, link_mensagem, titulo, vencimento):
    mensagem = (f"Olá!\n"
                f"{titulo}\n"
                f"{vencimento}\n"
                f"{link_mensagem}\n"
                f"Obrigado,\n"
                f"R.POSTAL SERVICOS CONTABEIS LTDA\n"
                f"veigaepostal@veigaepostal.com.br\n"
                f"(19)3829-8959")
    
    _abrir_chrome('https://web.whatsapp.com/', tela_inicial_site='nova_conversa.png', fechar_janela=False, outra_janela='email.png')
    _click_img('nova_conversa.png', conf=0.9)
    time.sleep(1)
    p.write(numero)
    time.sleep(2)

    if _find_img('sem_contato.png', conf=0.9):
        time.sleep(1)
        p.hotkey('ctrl', 'w')
        time.sleep(1)
        return 'Nenhum contato encontrado'

    while _find_img('procurando_numero.png', conf=0.9):
        time.sleep(1)

    p.press('enter')
    _wait_img('anexar.png', conf=0.9)
    time.sleep(1)
    
    while True:
        try:
            pyperclip.copy(mensagem)
            pyperclip.copy(mensagem)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass

    time.sleep(1)
    p.press('enter')
    time.sleep(3)
    
    while _find_img('anexar.png'):
        p.hotkey('ctrl', 'w')
        time.sleep(1)
        if _find_img('sair_do_site.png', conf=0.9):
            p.press('right')
            p.press('enter')
        time.sleep(3)
        
    
    time.sleep(1)
    return 'ok'
 
    
def enviar_anexo(numero, anexo, corpo_email):
    mensagem = (f"Olá!\n"
                f"{corpo_email}\n"
                f"Obrigado,\n"
                f"R.POSTAL SERVICOS CONTABEIS LTDA\n"
                f"veigaepostal@veigaepostal.com.br\n"
                f"(19)3829-8959")
    
    _abrir_chrome('https://web.whatsapp.com/', tela_inicial_site='nova_conversa.png', fechar_janela=False, outra_janela='email.png')

    _click_img('nova_conversa.png', conf=0.9)
    time.sleep(1)
    p.write(numero)
    time.sleep(2)

    if _find_img('sem_contato.png', conf=0.9):
        time.sleep(1)
        p.hotkey('ctrl', 'w')
        time.sleep(1)
        return 'Nenhum contato encontrado'

    while _find_img('procurando_numero.png', conf=0.9):
        time.sleep(1)

    p.press('enter')
    _wait_img('anexar.png', conf=0.9)
    _click_img('anexar.png', conf=0.9)
    time.sleep(1)
    _wait_img('documento.png', conf=0.9)
    _click_img('documento.png', conf=0.9)
    _wait_img('abrir.png', conf=0.9)
    time.sleep(1)
    
    while True:
        try:
            pyperclip.copy(anexo)
            pyperclip.copy(anexo)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass

    time.sleep(1)
    p.press('enter')
    _wait_img('digitar_mensagem.png', conf=0.9)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    while True:
        try:
            pyperclip.copy(mensagem)
            pyperclip.copy(mensagem)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
        
    time.sleep(1)
    p.press('enter')
    time.sleep(3)
    
    while p.locateOnScreen(r'imgs/aguardando.png'):
        time.sleep(1)
    
    time.sleep(1)
    p.hotkey('ctrl', 'w')
    time.sleep(1)
    
    return 'ok'


def mover_email(driver, pasta=''):
    # Volte para o contexto padrão (página principal)
    driver.switch_to.default_content()
    # clica no menu
    _click_img('menu_email.png', conf=0.9, clicks=2)
    
    # espera o menu abrir e clica na opção de mover
    _wait_img('mover_email.png', conf=0.9)
    time.sleep(0.5)
    _click_img('mover_email.png', conf=0.9, clicks=2)
    
    # aguarda a tela para selecionar e depois clica para abrir a lista de caixas
    _wait_img('tela_mover_email.png', conf=0.9)
    time.sleep(0.5)
    _click_img('tela_mover_email.png', conf=0.9)
    time.sleep(0.5)
    _wait_img('lista_caixas_email.png', conf=0.9)
    time.sleep(1)
    _click_img('lista_caixas_email.png', conf=0.9)
    
    time.sleep(1)
    # enquanto a caixa não aparecer aperta para descer a lista
    contador = 0
    while not _find_img(f'caixa_{pasta}_email.png', conf=0.9):
        p.press('down')
        contador += 1
        if contador > 10:
            if _find_img('inbox_1.png', conf=0.9):
                _click_img('inbox_1.png', conf=0.9)
            if _find_img('inbox_2.png', conf=0.9):
                _click_img('inbox_2.png', conf=0.9)
            if _find_img('tela_mover_email.png', conf=0.9):
                _click_img('tela_mover_email.png', conf=0.9)
                time.sleep(0.5)
                _wait_img('lista_caixas_email.png', conf=0.9)
                time.sleep(1)
                _click_img('lista_caixas_email.png', conf=0.9)
            contador = 0
    
    # seleciona a caixa
    _click_img(f'caixa_{pasta}_email.png', conf=0.9)
    
    while not _find_img('confirma_mover_email.png', conf=0.9):
        time.sleep(1)
        
    # confirmar a seleção
    _click_img('confirma_mover_email.png', conf=0.9)
    time.sleep(2)

    return driver


@_time_execution
@_barra_de_status
def run(window):
    print('>>> Aguardando documentos...')
    
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    # options.add_argument('--window-size=1920,1080')
    options.add_argument("--start-maximized")
    options.add_experimental_option('prefs', {
        "download.default_directory": "V:\\Setor Robô\\Scripts Python\\Geral\\Envia Documento\\ignore\\Anexos",  # Change default directory for downloads
        "download.prompt_for_download": False,  # To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True  # It will not show PDF directly in chrome
    })

    status, driver = _initialize_chrome(options)
    driver = login_email(driver)

    while True:
        mes = datetime.datetime.now().month
        ano = datetime.datetime.now().year
        nome_planilha = f'Envia link {mes}-{ano}'
        
        try:
            for arquivo in os.listdir(os.path.join('ignore', 'controle')):
                os.remove(os.path.join('ignore', 'controle', arquivo))
        except:
            pass
        
        print('>>> Aguardando e-mail')
        timer = 0
        while re.compile(r'<div class=\"no-items\" style=\"\">Sem itens para mostrar</div>').search(driver.page_source):
            time.sleep(1)
            timer += 1
            if timer > 60:
                driver.refresh()
                time.sleep(1)
                _wait_img('menu_classificacao.png', conf=0.9)
                _click_img('menu_classificacao.png', conf=0.9)
                _wait_img('crescente.png', conf=0.9)
                time.sleep(0.5)
                _click_img('crescente.png', conf=0.9)
                time.sleep(1)
                timer = 0
        
        nao_envia = 'x'
        # try:
        driver, titulo, cnpj, cnpj_limpo, corpo_email, vencimento, link_mensagem = captura_dados_email(driver)

        time.sleep(1)
        
        if titulo[:3] == 'Re:':
            time.sleep(1)
            driver = mover_email(driver, 'nao_enviados')
            _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo}', nome=nome_planilha, local=e_dir)
            print(f'{titulo}\n')
        else:
            if cnpj_limpo != 'x':
                cnpjs_iguais = verifica_o_numero(cnpj_limpo)
            else:
                cnpjs_iguais = False
            
            verifica_titulo = titulo.split('-')
            
            emails_diferentes = ['Desconsideração de e', 'Solicitação de documentos ', 'Documentos pendentes de visualização ']
            for diferente in emails_diferentes:
                if verifica_titulo[0] == diferente:
                    nao_envia = 'ok'
                    break
                else:
                    dias = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15',
                            '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']
                    for dia in dias:
                        if verifica_titulo[0] == f'Tarefa Boleto sistema de Gestão vencimento dia {dia} concluída ':
                            nao_envia = 'ok'
                            break
                        else:
                            nao_envia = False
                    break
                    
            titulo_sem_emoji = _remove_emojis(titulo)
            
            # se não tiver como enviar
            if nao_envia:
                time.sleep(1)
                driver = mover_email(driver, 'nao_enviados')
                _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo}', nome=nome_planilha + ' erros', local=e_dir)
                print(f'{titulo}\n')
            
            # se não encontrar o número da planilha
            elif not cnpjs_iguais:
                time.sleep(1)
                driver = mover_email(driver, 'nao_enviados')
                if cnpj_limpo == 'x':
                    try:
                        _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo};CNPJ não encontrado no corpo do e-mail', nome=nome_planilha + ' erros', local=e_dir)
                    except:
                        _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo_sem_emoji};CNPJ não encontrado no corpo do e-mail', nome=nome_planilha + ' erros', local=e_dir)
                    print('❌ CNPJ não encontrado no corpo do e-mail\n')
                else:
                    try:
                        _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo};Número não encontrado na planilha de dados', nome=nome_planilha + ' erros', local=e_dir)
                    except:
                        _escreve_relatorio_csv(f'{cnpj_limpo};x;x;{titulo_sem_emoji};Número não encontrado na planilha de dados', nome=nome_planilha + ' erros', local=e_dir)
                    print('❌ Número não encontrado na planilha de dados\n')
            
            # se tiver como enviar e encontrar o número na planilha
            else:
                resultado = ''
                for contato in cnpjs_iguais:
                    resultado = envia(resultado, contato, titulo, vencimento, link_mensagem, corpo_email, nome_planilha)
                # se der erro ao enviar
                if resultado == 'erro':
                    time.sleep(1)
                    driver = mover_email(driver, 'nao_enviados')
                    
                # se conseguir enviar
                elif resultado == 'ok':
                    time.sleep(1)
                    driver = mover_email(driver, 'enviados')
                # se der algum erro específico ao enviar
                else:
                    time.sleep(1)
                    driver = mover_email(driver, 'nao_enviados')
                
            """# se der erro em qualquer etapa
            except:
                time.sleep(1)
                print(driver.page_source)
                try:
                    _escreve_relatorio_csv(f'x;x;x;{titulo};Erro geral ao enviar a mensagem', nome=nome_planilha + ' erros', local=e_dir)
                except:
                    _escreve_relatorio_csv(f'x;x;x;{titulo_sem_emoji};Erro geral ao enviar a mensagem', nome=nome_planilha + ' erros', local=e_dir)
                print('❌ Erro ao enviar a mensagem\n')"""

        driver.refresh()
        time.sleep(1)
        _wait_img('menu_classificacao.png', conf=0.9)
        _click_img('menu_classificacao.png', conf=0.9)
        _wait_img('crescente.png', conf=0.9)
        time.sleep(0.5)
        _click_img('crescente.png', conf=0.9)
        time.sleep(1)
        

if __name__ == '__main__':
    run()
