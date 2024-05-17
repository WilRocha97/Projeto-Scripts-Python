import random, fitz, re, os, pyperclip, datetime
import pyautogui as p
import time

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def consulta_rest(empresa, caminho):
    cpf, nome, data_nascimento_sep = empresa
    data_nascimento = data_nascimento_sep.replace('/', '')

    # aguarda o cabeçalho do site
    while not _find_img('cabecalho_site.png', conf=0.9):
        if _find_img('bloquear_popup.png', conf=0.9):
            _click_img('bloquear_popup.png', conf=0.9)
        time.sleep(random.randrange(1, 5))

    print('>>> Preenchendo dados')
    _click_img('informe_cpf.png', conf=0.9)
    time.sleep(1)
    p.write(cpf)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write(data_nascimento)
    time.sleep(1)

    _click_img('sou_humano.png', conf=0.9)

    while not _find_img('sou_humano_check.png', conf=0.95):
        time.sleep(1)
        p.click()
        time.sleep(random.randrange(1, 5))

    _click_img('consultar.png', conf=0.9)

    print('>>> Aguardando consulta')
    while not _find_img('visualizar_restituicao.png', conf=0.9):
        for alerta in [('sem_dados.png', 'Não há informação para o exercício informado.'),
                       ('nao_confere.png', 'Data de nascimento não corresponde ao CPF informado.')]:
            if _find_img(alerta[0], conf=0.9):
                _click_img('sem_dados_fechar.png', conf=0.9)
                time.sleep(1)
                p.press('f5')
                return f'{cpf};{nome};{data_nascimento_sep};{alerta[1]}'
            time.sleep(1)

    arquivo_nome = f'Restituição - {nome}'
    
    salva_pdf_pagina(caminho, arquivo_nome)
    resposta = analise_consulta(caminho, arquivo_nome)
    
    _click_img('visualizar_restituicao.png', conf=0.9)
    while not _find_img('cabecalho_site.png', conf=0.9):
        time.sleep(random.randrange(1, 5))
    
    p.press('f5')
    
    return f'{cpf};{nome};{data_nascimento_sep};{resposta}'


def salva_pdf_pagina(caminho, arquivo_nome):
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
    
    while not _find_img('paisagem.png', conf=0.9):
        _click_img('retrato.png', conf=0.9, timeout=1)
        while not _find_img('paisagem_opcao.png', conf=0.9):
            time.sleep(1)
            p.click()
            time.sleep(1)
        _click_img('paisagem_opcao.png', conf=0.9, timeout=1)
    
    print('>>> Aguardando botão para salvar')
    # aguarda aparecer o botão de salvar e clica nele
    while not _find_img('botao_salvar.png', conf=0.9):
        if _find_img('botao_salvar_2.png', conf=0.9):
            break
    
    _click_img('botao_salvar.png', conf=0.9, timeout=1)
    _click_img('botao_salvar_2.png', conf=0.9, timeout=1)
    
    salva_arquivo(caminho, arquivo_nome, 'Consulta Restituição IRPF')


def salva_arquivo(caminho, arquivo_nome, arquivo):
    os.makedirs(caminho, exist_ok=True)
    
    print(f'>>> Salvando {arquivo.lower()}')
    
    # aguarda a tela de salvar do navegador abrir
    timer = 0
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_salvar_pdf_completo.png', conf=0.9):
            _click_img('salvar_como_pdf.png', conf=0.9, timeout=1)
            _click_img('salvar_como_pdf_2.png', conf=0.9, timeout=1)
        time.sleep(1)
        timer += 1
        if timer > 10:
            _click_img('salvar_como_pdf.png', conf=0.9, timeout=1)
            _click_img('salvar_como_pdf_2.png', conf=0.9, timeout=1)
    
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


def analise_consulta(caminho, arquivo_nome):
    situacao = 'Erro ao consultar a situação'
    arq = os.path.join(caminho, arquivo_nome + '.pdf')
    with fitz.open(arq) as pdf:
        # Para cada página do pdf
        for count, page in enumerate(pdf):
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)
            for regex in [r'seguinte situação:\n(.+)', r'Resultado encontrado: (.+)', r'2024\n(.+)']:
                situacao = re.compile(regex).search(textinho)
                if situacao:
                    situacao = situacao.group(1)
                    break

            if not situacao:
                print(textinho)
                raise Exception('Erro ao buscar com regex')
            print(situacao)

    return situacao.replace('\ufb01', 'fi')
    

@_time_execution
@_barra_de_status
def run(window):
    andamentos = f'Consulta Restituição IRPF'
    caminho = 'V:\Setor Robô\Scripts Python\Receita Federal\Consulta Restituição IRPF\execução\Consultas'
    
    _abrir_chrome('https://www.restituicao.receita.fazenda.gov.br/#/', iniciar_com_icone=True)
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        time.sleep(random.randrange(1, 5))
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
        
        resposta = consulta_rest(empresa, caminho)
        
        _escreve_relatorio_csv(resposta, nome=andamentos)
        
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    # função para abrir a lista de dados
    empresas = _open_lista_dados()

    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
