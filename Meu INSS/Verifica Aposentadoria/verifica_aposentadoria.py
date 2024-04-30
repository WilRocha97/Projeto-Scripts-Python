import datetime, re, time, os, pyperclip, pyautogui as p
from bs4 import BeautifulSoup

from sys import path
path.append(r'..\..\_comum')
from meu_inss_comum import _login_gov
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def busca_analise():
    _acessar_site_chrome('https://meu.inss.gov.br/#/simulador')
    print('>>> Aguardando simulação')
    # aguarda a tela da simulação carregar
    while not _find_img('tela_simulacao.png', conf=0.9):
        if _find_img('erro_requisicao_simulacao.png', conf=0.9):
            _acessar_site_chrome('https://meu.inss.gov.br/#/simulador')
            
        if _find_img('nao_simula_aposentadoria.png', conf=0.9):
            return False, 'A simulação de aposentadoria para quem só possui contribuições depois de 13/11/2019 está em desenvolvimento'
        time.sleep(1)

    return True, ''


def salva_pagina():
    p.hotkey('ctrl', 's')
    caminho = 'V:\Setor Robô\Scripts Python\Meu INSS\Verifica Aposentadoria\ignore\Página'
    os.makedirs('ignore\Página', exist_ok=True)
    arquivo_nome = 'meu_inss_pagina'
    salva_arquivo(caminho, arquivo_nome, 'Página')


def salva_arquivo(caminho, arquivo_nome, arquivo):
    print(f'>>> Salvando {arquivo.lower()}')
    
    # aguarda a tela de salvar do navegador abrir
    timer = 0
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_salvar_pdf_completo.png', conf=0.9):
            _click_img('salvar_pdf.png', conf=0.9, timeout=1)
            _click_img('salvar_pdf_2.png', conf=0.9, timeout=1)
        time.sleep(1)
        timer += 1
        if timer > 10:
            _click_img('salvar_pdf.png', conf=0.9, timeout=1)
            _click_img('salvar_pdf_2.png', conf=0.9, timeout=1)
            
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    pyperclip.copy(arquivo_nome)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy(caminho)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('substituir.png', conf=0.95):
        p.press('s')
    time.sleep(1)
    print(f'✔ {arquivo} baixado com sucesso')
    time.sleep(1)
    
    return True


def analisa_pagina(caminho, arquivo_nome):
    # Abra o arquivo e leia seu conteúdo
    while True:
        try:
            with open(os.path.join(caminho, arquivo_nome), 'r', encoding='utf-8') as arquivo:
                html_content = arquivo.read()
            break
        except:
            pass
        
    # Crie um objeto BeautifulSoup para analisar o HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # buscando todas as divs com uma classe específica:
    divs = soup.find_all('div', class_='MuiExpansionPanel-root')
    
    # Exiba os elementos encontrados
    aposentadoria_por = ''
    for div in divs:
        item = div.prettify()
        # print(item + '\n\n\n\n')
        if re.compile(r'(A partir dos dados simulados você\s+<strong>\s+tem\s+</strong>\s+direito ao benefício)').search(item):
            aposentadoria = re.compile(r'(Aposentadoria por.+)').search(item).group(1)
            if aposentadoria_por == '':
                aposentadoria_por = aposentadoria
            else:
                aposentadoria_por += f', {aposentadoria}'
    
    os.remove(os.path.join(caminho, arquivo_nome))
    
    if aposentadoria_por == '':
        return f'NÃO TEM direito ao benefício;'
    else:
        return f'TEM direito ao benefício;Aposentadoria por: {aposentadoria_por.replace("Aposentadoria por ", "")}'
        
    
def salva_pdf_pagina(caminho, data_hora_formatada, nome):
    _click_img('titulo_pagina.png', conf=0.9)
    p.mouseDown()
    
    p.press('end')
    
    x, y = p.locateCenterOnScreen('imgs\\rodape.png', confidence=0.9)
    p.moveTo(x-70, y-10)
    time.sleep(0.5)
    
    p.hotkey('ctrl', 'p')
    
    print('>>> Aguardando tela de impressão')
    _wait_img('tela_imprimir.png', conf=0.9)
    
    print('>>> Salvando PDF')
    imagens = ['print_to_pdf.png', 'print_to_pdf_2.png', 'print_to_pdf_3.png', 'print_to_pdf_4.png']
    for img in imagens:
        # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
        if _find_img(img, conf=0.9) or _find_img(img, conf=0.9):
            print('>>> Trocando destino da impressão')
            _click_img(img, conf=0.9, timeout=1)
            # aguarda aparecer a opção de salvar como PDF e clica nela
            while not _find_img('salvar_como_pdf.png', conf=0.9):
                if _find_img('salvar_como_pdf_2.png', conf=0.9):
                    break
            _click_img('salvar_como_pdf.png', conf=0.9, timeout=1)
            _click_img('salvar_como_pdf_2.png', conf=0.9, timeout=1)
    
    while not _find_img('definicoes.png', conf=0.9):
        _click_img('mais_definicoes.png', conf=0.9, timeout=1)
    
    time.sleep(1)
    p.press('pgDn')
    time.sleep(1)
    
    if not _find_img('apenas_selecao_marcado.png', conf=0.95):
        _click_img('apenas_selecao.png', conf=0.95)
    
    print('>>> Aguardando botão para salvar')
    # aguarda aparecer o botão de salvar e clica nele
    while not _find_img('botao_salvar.png', conf=0.9):
        if _find_img('botao_salvar_2.png', conf=0.9):
            break
            
    _click_img('botao_salvar.png', conf=0.9, timeout=1)
    _click_img('botao_salvar_2.png', conf=0.9, timeout=1)
    
    arquivo_nome = f'Simulação Aposentadoria {data_hora_formatada} - {nome} - Resumida'
    salva_arquivo(caminho, arquivo_nome, 'Simulação Aposentadoria Resumida')
    
    return True


@_time_execution
@_barra_de_status
def run(window):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        cpf, nome, senha = empresa

        andamentos = f'Verifica Aposentadoria'

        _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
        # faz login no cpf
        while True:
            resultado = _login_gov(cpf, senha)
            if resultado != 'recomeçar':
                break

        if resultado == 'ok':
            # busca o informe
            situacao, resultado = busca_analise()

            if situacao:
                p.hotkey('ctrl', 's')
                caminho = 'V:\Setor Robô\Scripts Python\Meu INSS\Verifica Aposentadoria\ignore\Página'
                os.makedirs('ignore\Página', exist_ok=True)
                arquivo_nome = 'meu_inss_pagina'
                resultado = salva_arquivo(caminho, arquivo_nome, 'Página')
                if resultado:
                    resposta_analise = analisa_pagina(caminho, arquivo_nome + '.html')
                else:
                    resposta_analise = 'Erro ao analisar a conteúdo do site;'
                    
                # Pega a data e hora atual
                data_hora_atual = datetime.datetime.now()
                # Formata a data e hora no formato desejado (dd-mm-aaaa hh-mm)
                formato = '%d-%m-%Y'
                data_hora_formatada = data_hora_atual.strftime(formato)
                caminho = f'V:\Setor Robô\Scripts Python\Meu INSS\Verifica Aposentadoria\execução\Simulações'
                os.makedirs(f'execução\Simulações', exist_ok=True)
                
                if salva_pdf_pagina(caminho, data_hora_formatada, nome):
                    analise_resumida = 'Analise resumida salvo com sucesso'
                else:
                    analise_resumida = 'Não salvou a analise resumida'
                    
                _click_img('baixar_pdf.png', conf=0.95)
                arquivo_nome = f'Simulação Aposentadoria {data_hora_formatada} - {nome} - Completa'
                if salva_arquivo(caminho, arquivo_nome, 'Simulação Aposentadoria Completa'):
                    analise_completa = 'Analise completa salvo com sucesso'
                else:
                    analise_completa = 'Não salvou a analise completa'
                resultado = f'{resposta_analise};{analise_resumida};{analise_completa}'

        _escreve_relatorio_csv(f'{cpf};{nome};{senha};{resultado}', nome=andamentos)
        p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
