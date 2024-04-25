# -*- coding: utf-8 -*-
import os, time, traceback, PySimpleGUI as sg, pandas as pd, pyautogui as p
import re
import time

from datetime import datetime
from openpyxl import load_workbook
from pandas import read_excel
from re import compile, search
from shutil import move
from threading import Thread


def escreve_doc(texto, local='Log', nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    for arq in os.listdir(local):
        try:
            os.remove(os.path.join(local, arq))
        except:
            pass
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()


def buscar_cpf_na_planilha(cpf, planilha_secundaria, coluna_cpf='CPF'):
    """
    Busca um CPF dentro de uma planilha.

    Args:
    cpf (str): O CPF que você deseja buscar.
    caminho_planilha (str): O caminho para a planilha (arquivo .xlsx ou .csv).
    coluna_cpf (str): O nome da coluna que contém os CPFs (padrão é 'CPF').

    Returns:
    pd.DataFrame: Um DataFrame contendo as linhas que correspondem ao CPF.
    """
    # Carregar a planilha
    try:
        df = pd.read_excel(planilha_secundaria)  # Para planilhas Excel (.xlsx)
        # Ou df = pd.read_csv(caminho_planilha)  # Para planilhas CSV
    except Exception as e:
        print(f"Erro ao carregar a planilha: {e}")
        return None

    # Verificar se a coluna CPF existe na planilha
    if coluna_cpf not in df.columns:
        return 'erro', f"A coluna '{coluna_cpf}' não foi encontrada na planilha secundária."

    # Converter a coluna de CPF para string (caso não esteja)
    df[coluna_cpf] = df[coluna_cpf].astype(str).str.strip()

    # Limpar o CPF de possíveis espaços em branco
    cpf = str(cpf).strip()

    # Verificar se algum CPF na coluna é igual ao CPF procurado
    resultado = df[df[coluna_cpf] == cpf]

    # Verificar se algum resultado foi encontrado
    if resultado.empty:
        print(f"O CPF {cpf} não foi encontrado na planilha.")
        return False, ''
    
    else:
        print(f"CPF {cpf} encontrado.")
        return 'ok', ''


def open_lista_dados(input_excel):
    if not input_excel:
        return False
    
    workbook = load_workbook(input_excel)
    workbook = workbook['Plan1']
    tipo_dados = 'xlsx'
    
    """# abre um alerta se não conseguir abrir o arquivo
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False"""
    
    return input_excel, workbook, tipo_dados


def run(window, planilha_primaria, planilha_secundaria):
    cpfs = pd.read_excel(planilha_primaria, index_col=None)
    encontrados = pd.DataFrame([[]])
    last_column = cpfs.shape[1]
    
    for count, cpfs_linha in enumerate(cpfs.itertuples(), start=1):
        cpf = cpfs_linha[last_column]
        print(cpf)
        
        if len(str(cpf)) > 11 or type(cpf) is not int:
            p.alert('CPF inválido na planilha primária.\n\nCertifique-se de que a última coluna da planilha seja a que contenha CPFs e que eles tenham no máximo 11 dígitos.')
            return
        
        resultado, mensagem =  buscar_cpf_na_planilha(cpf, planilha_secundaria)
        if resultado == 'ok':
            print(cpfs_linha)
            dataframe_temporario = pd.DataFrame([cpfs_linha])
            
            # Usar concat para adicionar a nova linha ao DataFrame original
            encontrados = pd.concat([encontrados, dataframe_temporario], ignore_index=True)
        
        if resultado == 'erro':
            p.alert(f"A coluna 'CPF' não foi encontrada na planilha secundária.\n\nCertifique-se de que exista uma coluna na planilha secundária com o nome 'CPF' e que eles tenham no máximo 11 dígitos.")
            return
            
        # atualiza a barra de progresso
        window['-progressbar-'].update_bar(count, max=int(len(cpfs)))
        window['-Progresso_texto-'].update(str(round(float(count) / int(len(cpfs)) * 100, 1)) + '%')
        window.refresh()
        
        # Verifica se o usuário solicitou o encerramento do script
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            p.alert(text='Busca encerrada.')
            return
    
    df = pd.DataFrame(encontrados).drop("Index", axis=1)
    
    # Pega a data e hora atual
    data_hora_atual = datetime.now()
    
    # Formata a data e hora no formato desejado (dd-mm-aaaa hh-mm)
    formato = '%d-%m-%Y %H-%M-%S'
    data_hora_formatada = data_hora_atual.strftime(formato)
    
    # Passo 2: Salve o DataFrame como um arquivo Excel
    os.makedirs('Resultados', exist_ok=True)
    caminho_arquivo = os.path.join('Resultados', f'CPFs econtrados {data_hora_formatada}.xlsx')  # Defina o caminho e o nome do arquivo Excel
    
    # Salve o DataFrame como um arquivo Excel
    df.to_excel(caminho_arquivo, index=False)
    
    p.alert(text=f'Planilha criada com sucesso: "CPFs econtrados {data_hora_formatada}.xlsx"')
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    # Layout da janela
    layout = [
        [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', disabled=True, border_width=0)],
        [sg.Text('')],
        [sg.Text('Selecione a planilha com os CPFs que você deseja consultar')],
        [sg.Text('(A ultima coluna da planilha deve ser a coluna com os CPFs):')],
        [sg.FileBrowse('Pesquisar', key='-abrir1-', file_types=(('Planilhas Excel', '*.xlsx'),)), sg.InputText(key='-input_excel_primaria-', size=80, disabled=True)],
        [sg.Text('Selecione a planilha onde a consulta será realizada')],
        [sg.Text('(A planilha deve ter uma coluna chamada "CPF" não imposta a posição):')],
        [sg.FileBrowse('Pesquisar', key='-abrir1-', file_types=(('Planilhas Excel', '*.xlsx'),)), sg.InputText(key='-input_excel_secundaria-', size=80, disabled=True)],
        [sg.Text('')],
        [sg.Text('', key='-Mensagens-')],
        [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color='#f0f0f0')],
        [sg.Button('Iniciar', key='-iniciar-', border_width=0), sg.Button('Encerrar', key='-encerrar-', disabled=True, border_width=0), sg.Button('Abrir resultados', key='-abrir_resultados-', border_width=0)],
    ]
    
    # guarda a janela na variável para manipula-la
    window = sg.Window('Busca CPF', layout)
    
    
    def run_script_thread():
        if not planilha_primaria:
            p.alert(text=f'Por favor insira a planilha com os CPFs que você deseja consultar.')
            return
        if not planilha_secundaria:
            p.alert(text=f'Por favor insira a planilha onde a consulta será realizada.')
            return
        
        # habilita e desabilita os botões conforme necessário
        window['-iniciar-'].update(disabled=True)
        window['-encerrar-'].update(disabled=False)
        
        window['-Mensagens-'].update('Gerando planilha resultado')
        # atualiza a barra de progresso para ela ficar mais visível
        window['-progressbar-'].update(bar_color=('#fca400', '#ffe0a6'))
        
        try:
            # Chama a função que executa o script
            run(window, planilha_primaria, planilha_secundaria)
        # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            traceback_str = traceback.format_exc()
            window['Log do sistema'].update(disabled=False)
            p.alert(text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
            escreve_doc(f'Traceback: {traceback_str}\n\n'
                        f'Erro: {erro}', local='Log')
                
        # habilita e desabilita os botões conforme necessário
        window['-iniciar-'].update(disabled=False)
        window['-encerrar-'].update(disabled=True)
        
        # apaga qualquer mensagem na interface
        window['-Mensagens-'].update('')
        # atualiza a barra de progresso para ela ficar mais visível
        window['-progressbar-'].update_bar(0)
        window['-progressbar-'].update(bar_color='#f0f0f0')
        window['-Progresso_texto-'].update('')
    
    
    categoria = None
    while True:
        # captura o evento e os valores armazenados na interface
        event, values = window.read()
        try:
            planilha_primaria = values['-input_excel_primaria-']
            planilha_secundaria = values['-input_excel_secundaria-']
        except:
            planilha_primaria = 'Desktop'
            planilha_secundaria = 'Desktop'
        
        if event == sg.WIN_CLOSED:
            break
        
        elif event == 'Log do sistema':
            os.startfile('Log')
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Busca CPF.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
        
        elif event == '-abrir_resultados-':
            os.makedirs('Resultados', exist_ok=True)
            os.startfile('Resultados')
    
    window.close()