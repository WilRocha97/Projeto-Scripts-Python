# -*- coding: utf-8 -*-
import shutil
import time
import fitz, re, os, atexit, sys, PySimpleGUI as sg
from threading import Thread
from functools import wraps
import PyPDF2
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime
from pathlib import Path
from pyautogui import alert

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

local_log='Cria E-book DIRPF\Log'
# Carregar a fonte TrueType (substitua 'sua_fonte.ttf' pelo caminho da sua fonte)
pdfmetrics.registerFont(TTFont('Fonte', 'V:\Setor Robô\Scripts Python\Analise de Arquivos\Cria E-book DIRPF\Assets\Montserrat-SemiBold.ttf'))


def chave_numerica(elemento):
    return int(elemento)


def escreve_doc(texto, local=local_log, nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    for arq in os.listdir(local):
        os.remove(arq)
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()


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


def escreve_relatorio_csv(texto, local, nome='Relatório', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local,  f"{nome} - complementar.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()


def ask_for_dir():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title='Selecione onde salvar a planilha',
    )
    
    return folder if folder else False


def decryption(input_name,output_name,password):
    pdfFile = open(input_name, "rb")
    reader = PyPDF2.PdfReader(pdfFile)
    writer = PyPDF2.PdfWriter()
    if reader.is_encrypted:
        reader.decrypt(password)
    for pageNum in range(len(reader.pages)):
        writer.add_page(reader.pages[pageNum])
    resultPdf = open(output_name, "wb")
    writer.write(resultPdf)
    pdfFile.close()
    resultPdf.close()


def quebra_de_linha(frase, caracteres_por_linha=30):
    palavras = frase.split()
    linhas = []
    linha_atual = ""

    for palavra in palavras:
        if len(linha_atual + palavra) <= caracteres_por_linha:
            linha_atual += palavra + " "
        else:
            linhas.append(linha_atual.strip())
            linha_atual = palavra + " "

    if linha_atual:
        linhas.append(linha_atual.strip())

    return linhas


def create_pdf(output_path, image_path, text):
    text = quebra_de_linha(text, caracteres_por_linha=17)
    
    # Abre o arquivo PDF para gravação
    pdf_canvas = canvas.Canvas(output_path, pagesize=(707,1007))

    # Adiciona a imagem à posição desejada
    pdf_canvas.drawInlineImage(image_path, 0, 0, width=707, height=1007)
    
    altura_linha = 401
    for linha in text:
        # Define as configurações de texto
        text_object = pdf_canvas.beginText(x=44, y=altura_linha,)
        altura_linha = (altura_linha - 35)
        
        text_object.setFont("Fonte", 30)
        text_object.setFillColor(colors.black)
    
        # Adiciona o texto personalizável
        text_object.textLine(linha.upper())
    
        # Desenha o texto no canvas
        pdf_canvas.drawText(text_object)
    
    # Fecha o arquivo PDF
    pdf_canvas.save()
    

def run(window, pasta_inicial, pasta_final):
    # inicia o dicionário
    subpasta_arquivos = {}

    window['-Mensagens-'].update(f'Analisando arquivos...')
    # itera sobre todas as subpastas dentro da pasta mestre
    for count, nome_subpasta in enumerate(os.listdir(pasta_inicial), start=1):
        caminho_subpasta = os.path.join(pasta_inicial, nome_subpasta)
        
        # Verifica se é uma pasta
        if os.path.isdir(caminho_subpasta):
            print(f"Entrando na subpasta: {nome_subpasta}")
            
            # para cada arquivo na subpasta
            for arquivo in os.listdir(caminho_subpasta):
                # verifica se o arquivo tem senha, se tiver tira a senha dele
                try:
                    with fitz.open(os.path.join(caminho_subpasta, arquivo)) as pdf:
                        arquivo_sem_senha = pdf
                except:
                    try:
                        password = re.compile(r'Senha.(\w+)').search(arquivo).group(1)
                        decryption(os.path.join(caminho_subpasta, arquivo), os.path.join(caminho_subpasta, arquivo), password)
                    except:
                        # se tiver senha, mas a senha não for informada no nome do arquivo, pula para próxima subpasta
                        alert(f'Não é possível criar E-Book de {nome_subpasta}.\n'
                                f'Existe PDF protegido sem a senha informada no nome do arquivo.\n'
                                f'Arquivo protegido encontrado: {os.path.join(caminho_subpasta, arquivo)}'
                                f'Para que o processo automatizado mescle PDF protegido, a senha deve ser informada no nome do arquivo, por exemplo:\n'
                                f'Senha 1234567.pdf\n')
                        break
                
                # se a subpasta já consta no dicionário, adiciona mais um pdf
                if nome_subpasta in subpasta_arquivos:
                    subpasta_arquivos[nome_subpasta].append(arquivo)
                # se não cria uma nova subpasta dentro do dicionário já adicionando o pdf
                else:
                    subpasta_arquivos[nome_subpasta] = [arquivo]
                    
        window['-progressbar-'].update_bar(count, max=int(len(os.listdir(pasta_inicial))))
        window['-Progresso_texto-'].update(str(round(float(count) / int(len(os.listdir(pasta_inicial))) * 100, 1)) + '%')
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
            
        window.refresh()
     
    # ---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    os.makedirs('Arquivos para mesclar', exist_ok=True)
    # limpa a pasta de cópias de arquivos
    for arquivo in os.listdir('Arquivos para mesclar'):
        os.remove(os.path.join('Arquivos para mesclar', arquivo))
        
    window['-Mensagens-'].update(f'Criando arquivos...')
    window['-progressbar-'].update_bar(0)
    window['-Progresso_texto-'].update('')
    
    # para cada subpasta
    for count, (subpasta, nomes_arquivos) in enumerate(subpasta_arquivos.items(), start=1):
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
        
        contador = 3
        lista_arquivos = []
        # cria uma cópia numerada dos PDFs
        for arquivo in nomes_arquivos:
            if re.compile(r'imagem-recibo').search(arquivo):
                # busca o nome do declarante
                with fitz.open(os.path.join(pasta_inicial, subpasta, arquivo)) as pdf:
                    for page in pdf:
                        texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                        nome_declarante = re.compile(r'Nome do declarante\n.+\n(.+)').search(texto_pagina).group(1)
                        break
                        
                # cria a capa
                nome_do_arquivo = os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf')
                caminho_da_imagem = "Assets\DIRPF_capa.png"
                create_pdf(nome_do_arquivo, caminho_da_imagem, nome_declarante)
                
                # adiciona o arquivo da capa na lista da subpasta
                nomes_arquivos.append(f'Capa E-book {nome_declarante}.pdf')
                shutil.copy(os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
                lista_arquivos.append(0)
                
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '1.pdf'))
                lista_arquivos.append(1)
                continue
                
            if re.compile(r'imagem-declaracao').search(arquivo):
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
                lista_arquivos.append(2)
                continue
            if re.compile(r'INFORME').search(arquivo.upper()):
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                lista_arquivos.append(contador)
                contador += 1
                continue
        
        count = 0
        while not _find_img('pessoal.png', conf=0.9):
            print('oizzzz')
            
            if _find_img('pessoal1.png', conf=0.9):
                print('break')
                break
                
            if _find_img('nao_existe.png', conf=0.9):
                return driver, 'Não Existe Pastas', 'ok'
    
            time.sleep(1)
            count += 1
            if count > 10:
                return driver, 'Cliente não cadastrado Onvio', 'ok'
        
        for arquivo in nomes_arquivos:
            if re.compile(r'Capa E-book').search(arquivo):
                continue
            if re.compile(r'imagem-recibo').search(arquivo):
                continue
            if re.compile(r'imagem-declaracao').search(arquivo):
                continue
            if re.compile(r'INFORME').search(arquivo.upper()):
                continue
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                lista_arquivos.append(contador)
                contador += 1
                
        # ordena a lista de arquivos
        lista_arquivos = sorted(lista_arquivos, key=chave_numerica)
        
        # mescla os arquivos
        pdf_merger = PyPDF2.PdfMerger()
        for arquivo in lista_arquivos:
            caminho_completo = os.path.join('Arquivos para mesclar', f'{arquivo}.pdf')
            pdf_merger.append(caminho_completo)
        
        # cria o e-book
        unificado_pdf = os.path.join(pasta_final, f'E-BOOK DIRPF - {nome_declarante}.pdf')
        pdf_merger.write(unificado_pdf)
        pdf_merger.close()
        
        window['-progressbar-'].update_bar(count, max=int(len(subpasta_arquivos.items())))
        window['-Progresso_texto-'].update(str(round(float(count) / int(len(subpasta_arquivos.items())) * 100, 1)) + '%')
        window.refresh()
        
        # limpa a pasta de cópias de arquivos
        for arquivo in os.listdir('Arquivos para mesclar'):
            os.remove(os.path.join('Arquivos para mesclar', arquivo))
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
                    
    alert(text='PDFs unificados com sucesso.')
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    # Especifique o caminho para o arquivo de trava
    lock_file_path = 'integra_contador.lock'
    
    # Verifique se outra instância está em execução
    if not create_lock_file(lock_file_path):
        alert(text="Outra instância já está em execução.")
        sys.exit(1)
    
    # Defina uma função para remover o arquivo de trava ao final da execução
    atexit.register(remove_lock_file, lock_file_path)
    
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    layout = [
        [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True)],
        [sg.Text('')],
        [sg.Text('Selecione a pasta que contenha as subpastas dos declarantes com os arquivos PDF para analisar:')],
        [sg.FolderBrowse('Pesquisar', key='-abrir_pdf-'), sg.InputText(key='-output_dir-', size=80, disabled=True)],
        [sg.Text('Selecione a pasta para salvar os e-books:')],
        [sg.FolderBrowse('Pesquisar', key='-abrir_pdf_final-'), sg.InputText(key='-output_dir-', size=80, disabled=True)],
        [sg.Text('')],
        [sg.Text('', key='-Mensagens-')],
        [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color='#f0f0f0')],
        [sg.Button('Iniciar', key='-iniciar-', border_width=0), sg.Button('Encerrar', key='-encerrar-', disabled=True, border_width=0), sg.Button('Abrir resultados', key='-abrir_resultados-', disabled=True, border_width=0)],
    ]
    
    # guarda a janela na variável para manipula-la
    window = sg.Window('Cria E-book DIRPF', layout)
    
    def run_script_thread():
        # try:
        if not pasta_inicial:
            alert(text=f'Por favor informe a pasta com arquivos PDF para analisar.')
            return
        if not pasta_final:
            alert(text=f'Por favor informe um diretório para salvar os arquivos unificados.')
            return
            
        # habilita e desabilita os botões conforme necessário
        window['-abrir_pdf-'].update(disabled=True)
        window['-abrir_pdf_final-'].update(disabled=True)
        window['-iniciar-'].update(disabled=True)
        window['-encerrar-'].update(disabled=False)
        window['-abrir_resultados-'].update(disabled=False)
        # apaga qualquer mensagem na interface
        window['-Mensagens-'].update('')
        # atualiza a barra de progresso para ela ficar mais visível
        window['-progressbar-'].update(bar_color=('#fca400', '#ffe0a6'))
        
        #try:
        # Chama a função que executa o script
        run(window, pasta_inicial, pasta_final)
        """# Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            escreve_doc(erro)
            window['Log do sistema'].update(disabled=False)"""
        
        window['-progressbar-'].update_bar(0)
        window['-progressbar-'].update(bar_color='#f0f0f0')
        window['-Progresso_texto-'].update('')
        window['-Mensagens-'].update('')
        # habilita e desabilita os botões conforme necessário
        window['-abrir_pdf-'].update(disabled=False)
        window['-abrir_pdf_final-'].update(disabled=False)
        window['-iniciar-'].update(disabled=False)
        window['-encerrar-'].update(disabled=True)
        
    
    while True:
        # captura o evento e os valores armazenados na interface
        event, values = window.read()
        
        try:
            pasta_inicial = values['-abrir_pdf-']
            pasta_final = values['-abrir_pdf_final-']
        except:
            pasta_inicial = None
            pasta_final = None
        
        if event == sg.WIN_CLOSED:
            break
        
        elif event == 'Log do sistema':
            os.startfile(os.path.join(local_log, 'Log.txt'))
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Cria E-book DIRPF.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
        
        elif event == '-abrir_resultados-':
            os.startfile(pasta_final)
    
    window.close()
    
    
    