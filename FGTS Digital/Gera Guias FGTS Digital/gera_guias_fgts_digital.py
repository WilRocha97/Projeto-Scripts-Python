import time
import os
import fitz
import re
import shutil
import pyperclip

from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img, _click_position_img, _get_comp
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def seleciona_certificado():
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('botao_gov.png', conf=0.9):
        time.sleep(1)
    
    # clica no botão do gov.br
    _click_img('botao_gov.png', conf=0.9)
    
    # aguarda o botão de certificado
    while not _find_img('certificado.png', conf=0.95):
        time.sleep(1)
        
    # clica no botão de certificado
    _click_img('certificado.png', conf=0.95)
    
    # aguarda a seleção de certificado
    timer = 0
    while not _find_img('seleciona_certificado.png'):
        if _find_img('captcha.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return 'recomeçar'
        time.sleep(1)
        timer += 1
        
        if timer > 10:
            _acessar_site_chrome('https://fgtsdigital.sistema.gov.br/portal/login')
            time.sleep(1)
            while not _find_img('botao_gov.png', conf=0.9):
                time.sleep(1)
            
            # clica no botão do gov.br
            _click_img('botao_gov.png', conf=0.9)
            
            # aguarda o botão de certificado
            while not _find_img('certificado.png', conf=0.95):
                if _find_img('erro_gov.png'):
                    p.press('pgDn')
                    time.sleep(1)
                    # clica no botão do gov.br
                    _click_img('botao_gov.png', conf=0.9)
                    
            # clica no botão de certificado
            _click_img('certificado.png', conf=0.95)
            timer = 0
            
    p.press('enter')


def login(cnpj):
    # aguarda a seleção de perfil
    while not _find_img('defina_perfil.png'):
        if _find_img('troca_perfil.png'):
            break
            
        if _find_img('trocar_perfil.png', conf=0.95):
            _click_img('trocar_perfil.png', conf=0.95)
            
        if _find_img('erro_gov.png'):
            p.press('pgDn')
            time.sleep(1)
            # clica no botão do gov.br
            _click_img('botao_gov.png', conf=0.9)
        time.sleep(1)
    
    if _find_img('aceita_cookies.png', conf=0.95):
        _click_img('aceita_cookies.png', conf=0.95)
    
    if not _find_img('procurador_selecionado.png', conf=0.95):
        # clica para definir perfil
        _click_img('seleciona_perfil.png', conf=0.9, timeout=1)
        
        # aguarda aparecer procurador
        while not _find_img('perfil_procurador.png'):
            if _find_img('seleciona_perfil_2.png', conf=0.95):
                _click_img('seleciona_perfil_2.png', conf=0.95, timeout=1)
                time.sleep(0.5)
                p.write('Procurador')
                time.sleep(0.5)
                p.press('enter')
                
            if _find_img('procurador_cinza.png'):
                break
            time.sleep(1)
            
        # clica em procurador
        _click_img('perfil_procurador.png', conf=0.9, timeout=1)
    
    time.sleep(1)
    
    # clica no campo
    _click_position_img('informe_cnpj.png', '+', pixels_y=39, clicks=3, conf=0.95)
    time.sleep(1)
    
    p.write(cnpj)
    time.sleep(1)
    
    # clica em definir ou selecionar
    _click_img('definir.png', conf=0.9, timeout=1)
    _click_img('selecionar.png', conf=0.9, timeout=1)
    time.sleep(1)
    
    if _find_img('sem_procuracao.png', conf=0.95):
        return 'Não existe procuração para o Número de Inscriçao selecionado'
    
    return 'ok'


def busca_guia(comp):
    # aguarda aparecer o campo de CNPJ
    while not _find_img('gestao_guias.png', conf=0.9):
        if _find_img('cadastro_incompleto.png', conf=0.9):
            return 'Cadastro incompleto'
        time.sleep(1)
    
    _acessar_site_chrome('https://fgtsdigital.sistema.gov.br/cobranca/#/gestao-guias/emissao-guia-rapida')
    
    print('>>> Buscando guia')
    while not _find_img('emissao_de_guia.png', conf=0.9):
        time.sleep(1)
        
    if _find_img('sem_debitos.png', conf=0.95):
        _click_img('sem_debitos.png', conf=0.95, clicks=3)
        while True:
            try:
                p.hotkey('ctrl', 'c')
                p.hotkey('ctrl', 'c')
                mensagem = pyperclip.paste()
                break
            except:
                pass
        
        mensagem = mensagem.split('(')
        mensagem = f'{mensagem[0]};{mensagem[1].replace(")", "")}- {mensagem[2].replace(")", "")}'
        
        return mensagem.replace("\n", "").replace("\r", "")
    
    # clica para editar a competência
    _click_position_img('competencia_apuracao.png', '+', pixels_y=57, conf=0.9)
    
    # digitar competência
    p.write(comp)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    
    # clica em pesquisar
    _click_img('pesquisar.png', conf=0.95, clicks=3)
    
    return 'Empresa ok'


def salva_guia(cnpj, nome):
    os.makedirs('execução\Guias', exist_ok=True)
    print('>>> Baixando Informe')
    caminho = 'V:\Setor Robô\Scripts Python\FGTS Digital\Gera Guias FGTS Digital\execução\Guias'
    
    arquivo_nome = f'GFD - {comp.replace("/", "-")} - Vencimento- - {cnpj} - {nome}'

    # desce a tela até achar o botão de download
    while not _find_img('emitir_guia.png', conf=0.9):
        p.press('down')

    # clica para baixar o pdf
    _click_img('emitir_guia.png', conf=0.9)

    # aguarda a tela de salvar do navegador abrir
    _wait_img('salvar_como.png', conf=0.9, timeout=-1)
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
    print('✔ Guia GFD gerada')
    time.sleep(5)
    analisa_guia(os.path.join(caminho, arquivo_nome + '.pdf'))

    return 'Guia gerada com sucesso'


def analisa_guia(guia):
    texto_arquivo = ''
    with fitz.open(guia) as pdf:
        for page in pdf:
            texto_pagina = page.get_text('text', flags=1 + 2 + 8)
            texto_arquivo += texto_pagina
    
    vencimento = re.compile(r'Pagar este documento até\n.+\n(\d\d/\d\d/\d\d\d\d)').search(texto_arquivo).group(1)
    vencimento = vencimento.replace('/', '-')
    shutil.move(guia, guia.replace('Vencimento-', f'Vencimento-{vencimento}'))


@_time_execution
@_barra_de_status
def run(window):
    andamentos = f'Guias FGTS Digital'
    total_empresas = empresas[index:]
    while True:
        _abrir_chrome('https://fgtsdigital.sistema.gov.br/portal/login')
        resultados = seleciona_certificado()
        if resultados != 'recomeçar':
            break
            
    for count, empresa in enumerate(empresas[index:], start=1):
        comeco = datetime.now()
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)
        
        cnpj, nome = empresa
        
        resultado = login(cnpj)
        
        if resultado == 'Não existe procuração para o Número de Inscriçao selecionado':
            _escreve_relatorio_csv(f'{cnpj};{nome};{resultado}', nome=andamentos)
            continue
                
        resultado = busca_guia(comp)
        
        if resultado == 'Empresa ok':
            # Salvar a guia
            resultado = salva_guia(cnpj, nome)
        
        _escreve_relatorio_csv(f'{cnpj};{nome};{resultado}', nome=andamentos)
        _click_img('botao_home.png', conf=0.9)
        
    time.sleep(2)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    comp = _get_comp(printable='mm/aaaa', strptime='%m/%Y')
    
    if index is not None:
        run()
