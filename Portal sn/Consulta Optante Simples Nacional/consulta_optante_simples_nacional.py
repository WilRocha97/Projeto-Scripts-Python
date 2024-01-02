import pyautogui as p
import time
import os
import pyperclip
from bs4 import BeautifulSoup
from threading import Thread
from pathlib import Path

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status, _escreve_header_csv


def limpa_info(info):
    info = info.replace(': ', ':')
    info_limpa = info.split(':')
    info = info_limpa[1]
    return info


def copiar(limpar_info=False, deletar_final=0):
    while True:
        try:
            p.hotkey('ctrl', 'c')
            p.hotkey('ctrl', 'c')
            texto = pyperclip.paste()
            if limpar_info:
                texto = limpa_info(texto)
            break
        except:
            pass
    time.sleep(0.2)
    p.click()
    return str(texto[:-deletar_final]).replace("/", "-").replace("\n", "-")


def login(cnpj):
    # espera o site abrir
    while not _wait_img('cnpj.png', conf=0.9, timeout=-1):
        time.sleep(1)
    
    _click_img('cnpj.png', conf=0.9)
    time.sleep(1)
    p.write(cnpj)
    p.press('enter')
    
    while not _wait_img('consulta.png', conf=0.9, timeout=-1):
        time.sleep(1)
    
    return True


def analise(cnpj, nome):
    _click_img('optante_sn.png', conf=0.9, clicks=3)
    situacao_sn = copiar(limpar_info=True, deletar_final=2)
    
    _click_img('optante_simei.png', conf=0.9, clicks=3)
    situacao_simei = copiar(limpar_info=True, deletar_final=2)
        
    _click_img('mais_info.png', conf=0.9)
    
    while not _find_img('periodos_anteriores.png', conf=0.9):
        time.sleep(1)
    
    if _find_img('situacao_anterior.png', conf=0.9):
        p.press('down', presses=3)
        
        _click_position_img('data_inicial.png', '+', pixels_y=37, conf=0.9, clicks=3)
        data_inicial_anterior = copiar(deletar_final=1)
        
        _click_position_img('data_final.png', '+', pixels_y=37, conf=0.9, clicks=3)
        data_final_anterior = copiar(deletar_final=1)
        
        _click_position_img('detalhamento.png', '+', pixels_y=37, conf=0.9, clicks=3)
        opcoes_sn = copiar(deletar_final=2)
        
        salvar_pdf(f'{cnpj} - {nome} - Excluída do Simples Nacional em {data_final_anterior}')
    
    else:
        data_inicial_anterior = ''
        data_final_anterior = ''
        
        while not _find_img('opcoes_sn.png', conf=0.9, ):
            p.press('down')
        _click_img('opcoes_sn.png', conf=0.9, clicks=3)
        opcoes_sn = copiar(limpar_info=True, deletar_final=2)
    
    while not _find_img('opcoes_simei.png', conf=0.9,):
        p.press('down')
    _click_img('opcoes_simei.png', conf=0.9, clicks=3)
    opcoes_simei = copiar(limpar_info=True, deletar_final=2)
        
    p.hotkey('ctrl', 'w')
    return situacao_sn, situacao_simei, opcoes_sn, opcoes_simei, data_inicial_anterior, data_final_anterior
    
 
def salvar_pdf(arquivo):
    download_folder = 'V:\Setor Robô\Scripts Python\Portal sn\Consulta Optante Simples Nacional\execução\Relatórios'
    p.hotkey('ctrl', 'p')
    
    print('>>> Aguardando tela de impressão')
    _wait_img('tela_imprimir.png', conf=0.9)
    
    print('>>> Salvando PDF')
    imagens = ['print_to_pdf.png', 'print_to_pdf_2.png']
    for img in imagens:
        # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
        if _find_img(img, conf=0.9) or _find_img(img, conf=0.9):
            _click_img(img, conf=0.9)
            # aguarda aparecer a opção de salvar como PDF e clica nela
            _wait_img('salvar_como_pdf.png', conf=0.9)
            _click_img('salvar_como_pdf.png', conf=0.9)
    
    # aguarda aparecer o botão de salvar e clica nele
    while not _find_img('botao_salvar.png', conf=0.9):
        time.sleep(1)
    
    p.click(p.locateCenterOnScreen(f'imgs/botao_salvar.png', confidence=0.9))
    
    while not _find_img('salvar_como.png', conf=0.9):
        print('>>> Salvando relatório')
        p.click(p.locateCenterOnScreen(f'imgs/botao_salvar.png', confidence=0.9))
        time.sleep(1)
    
    os.makedirs(download_folder, exist_ok=True)
    
    while True:
        try:
            pyperclip.copy(arquivo)
            time.sleep(1)
            pyperclip.copy(arquivo)
            time.sleep(1)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
        
    time.sleep(1)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    while True:
        try:
            pyperclip.copy(download_folder)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(1)
    

@_time_execution
@_barra_de_status
def run(window):
    andamentos = f'Consulta Optantes Simples Nacional'
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count + index)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index))} Restantes')
        _indice(count, total_empresas, empresa, index)
        
        cnpj, nome = empresa
        nome = nome.replace('/', '')
        
        _abrir_chrome('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes')
        if not login(cnpj):
            continue
        
        situacao_sn, situacao_simei, opcoes_sn, opcoes_simei, data_inicial_anterior, data_final_anterior = analise(cnpj, nome)
        _escreve_relatorio_csv(f'{cnpj};{nome};{situacao_sn};{situacao_simei};{opcoes_sn};{opcoes_simei};{data_inicial_anterior};{data_final_anterior}', nome=andamentos)
    
    _escreve_header_csv('CNPJ;NOME;OPTANTE SN;OPTANTE SIMEI;OCORRÊNCIAS SN;OCORRÊNCIAS SIMEI;DATA DE INÍCIO SN;DATA FINAL SN', nome=andamentos)

if __name__ == '__main__':
    run()
