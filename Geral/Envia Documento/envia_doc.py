# -*- coding: utf-8 -*-
import re, random, time, os, shutil, pyperclip, pyautogui as p, pandas as pd
from pathlib import Path

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img


def verifica_o_numero(arquivo):
    # Abrir o pdf
    '''with fitz.open(arquivo) as pdf:
        # Para cada página do pdf
        for page in pdf:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
            try:
                cnpj = re.compile(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)').search(textinho).group(1)
            except:
                try:
                    cnpj = re.compile(r'(\d\d\d.\d\d\d.\d\d\d-\d\d').search(textinho).group(1)
                except:
                    try:
                        cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)').search(textinho).group(1)
                    except:
                        try:
                            cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d').search(textinho).group(1)
                        except:
                            return False'''
    
    cnpj = '12345678912345'
    # Definir o nome das colunas
    colunas = ['cnpj', 'numero', 'nome']
    # Localiza a planilha
    caminho = Path('V:/Setor Robô/Scripts Python/Geral/WhatsApp/ignore/Dados.csv')
    # Abrir a planilha
    lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
    # Definir o index da planilha
    lista.set_index('cnpj', inplace=True)

    numero = lista.loc[int(cnpj)]
    numero = list(numero)
    numero = numero[0]
    
    return cnpj, numero


def envia(numero, arq):
    # abre o chat para o qual o PDF será enviado
    while not _find_img('pesquisar_contato.png', conf=0.9):
        print('Aguardando conversas do WhatsApp')
        time.sleep(1)
    
    # clica na barra de pesquisa
    _click_img('pesquisar_contato.png', conf=0.9)
    # espera a barra habilitar
    _wait_img('inserir_numero.png', conf=0.9)
    #  escreve o número
    p.write(numero)
    
    # espera o resultado da busca aparecer
    while not _find_img('conversas.png', conf=0.8):
        time.sleep(1)
        # se não aparecer nenhum contato retorna
        if _find_img('sem_contato.png', conf=0.9):
            return 'Contato não encontrado'
    
    # clica no contato para abrir a conversa
    _click_position_img('conversas.png', '+', pixels_y=70, conf=0.8)
    time.sleep(3)
    
    print('Anexando arquivo')
    # clica para anexar item
    _click_img('anexar.png', conf=0.9)
    
    # espera as opções aparecerem
    _wait_img('doc.png', conf=0.9, timeout=1)
    #clica para anexar documento
    _click_img('doc.png', conf=0.9)
    
    # seleciona o arquivo
    seleciona_arquivo(arq)
    
    #espera a tela para inserir uma mensagem com o documento
    _wait_img('mensagem.png', conf=0.9)
    
    # escreve a mensagem e envia
    p.write('Segue um novo documento do escritório Veiga & Postal')
    time.sleep(0.5)
    p.press('enter')
    
    return True

    
def seleciona_arquivo(arq):
    _wait_img('abrir.png', conf=0.9)
    
    p.press('tab', presses=5, interval=0.1)
    p.press('enter')
    time.sleep(0.5)
    
    pyperclip.copy('V:\Setor Robô\Scripts Python\Geral\WhatsApp\ignore\arquivos')
    pyperclip.paste()
    time.sleep(0.5)
    
    p.press('enter')
    time.sleep(0.5)
    
    p.press('tab', presses=6, interval=0.1)
    p.press('enter')
    time.sleep(0.5)
    
    p.write(arq)
    time.sleep(0.5)
    
    p.hotkey('alt', 'a')
    

@_time_execution
def run():
    download_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\ignore\\arquivos"
    final_folder_enviado = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\Execução\\Enviados"
    final_folder_nao_enviado = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\Execução\\Não Enviados"
    while 0 < 1:
        print('>>> Aguardando documentos...')
        # limpa a pasta de download caso fique algum arquivo nela
        for arq in os.listdir(download_folder):
            # determina o tempo de espera entre uma mensagem e outra para tentar evitar span
            numero = random.randint(1, 10)
            time.sleep(numero)
            
            arquivo = os.path.join(download_folder, arq)
            
            cnpj, numero = verifica_o_numero(arquivo)
            print(numero)
            
            if not numero:
                os.makedirs(final_folder_nao_enviado, exist_ok=True)
                time.sleep(1)
                shutil.move(arquivo, os.path.join(final_folder_nao_enviado, arq))
                _escreve_relatorio_csv(f'{cnpj};{numero};CNPJ não encontrado na planilha', nome='Envia guia')
                
            else:
                resultado = envia(numero, arq)
                if resultado == 'Contato não encontrado':
                    os.makedirs(final_folder_nao_enviado, exist_ok=True)
                    time.sleep(1)
                    shutil.move(arquivo, os.path.join(final_folder_nao_enviado, arq))
                    _escreve_relatorio_csv(f'{cnpj};{numero};Contato não encontrado', nome='Envia guia')
                else:
                    os.makedirs(final_folder_enviado, exist_ok=True)
                    time.sleep(1)
                    shutil.move(arquivo, os.path.join(final_folder_enviado, arq))
                    _escreve_relatorio_csv(f'{cnpj};{numero};Arquivo enviado', nome='Envia guia')
    

if __name__ == '__main__':
    run()

