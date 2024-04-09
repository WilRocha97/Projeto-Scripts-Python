# -*- coding: utf-8 -*-
import shutil, io, fitz, PyPDF2, re, os, sys, PySimpleGUI as sg
import time

from PIL import Image
from threading import Thread
from pyautogui import alert
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Carregar a fonte TrueType (substitua 'sua_fonte.ttf' pelo caminho da sua fonte)
pdfmetrics.registerFont(TTFont('Fonte', 'Assets\Montserrat-SemiBold.ttf'))
pdfmetrics.registerFont(TTFont('Tabela', 'Assets\JetBrainsMono-VariableFont_wght.ttf'))

def chave_numerica(elemento):
    return int(elemento)


def escreve_doc(texto, local='Log', nome='Log', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    for arq in os.listdir(local):
        os.remove(os.path.join(local, arq))
    
    try:
        f = open(os.path.join(local, f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(texto))
    f.close()


def escreve_relatorio_csv(texto, local, nome='Relatório', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local,  f"{nome} - complementar.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()


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


def create_pdf(output_path, image_path, text, text_2):
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
    
    # adiciona o segundo texto
    # Define as configurações de texto
    text_object = pdf_canvas.beginText(x=44, y=296, )
    
    text_object.setFont("Fonte", 15)
    text_object.setFillColor(colors.black)
    
    # Adiciona o texto personalizável
    text_object.textLine(text_2.upper())
    
    # Desenha o texto no canvas
    pdf_canvas.drawText(text_object)
    
    # Fecha o arquivo PDF
    pdf_canvas.save()
    

def cria_pagina_resumo(infos_resumo, output_path):
    # c.rect(x, y, width, height, fill=1)  #draw rectangle
    
    infos_resumo = sorted(infos_resumo)
    # Abre o arquivo PDF para gravação
    pdf_canvas = canvas.Canvas(output_path, pagesize=(707, 1007))
    
    altura_linha = 910
    """pdf_canvas.setFillColorRGB(255, 100, 0)
    pdf_canvas.rect(0, altura_linha-35, 707, 70, fill=1, stroke=0)"""
    
    for info in infos_resumo:
        try:
            infos = info[1].upper().split(':')
            
            tamanho = 85 - len(info[1])
            info = f"{infos[0]} {'.' * tamanho}{infos[1]}"
            print(info)
            # Define as configurações de texto
            text_object = pdf_canvas.beginText(x=100, y=altura_linha, )
            altura_linha = (altura_linha - 35)
            text_object.setFont("Tabela", 10)
            text_object.setFillColor(colors.black)
            # Adiciona o texto personalizável
            text_object.textLine(info)
            # Desenha o texto no canvas
            pdf_canvas.drawText(text_object)
        except:
            info = info[1]
            text = quebra_de_linha(info, caracteres_por_linha=36)
            for titulo in text:
                # Define as configurações de texto
                text_object = pdf_canvas.beginText(x=100, y=altura_linha, )
                altura_linha = (altura_linha - 17)
                text_object.setFont("Fonte", 15)
                text_object.setFillColor(colors.black)
                # Adiciona o texto personalizável
                text_object.textLine(titulo.upper())
                # Desenha o texto no canvas
                pdf_canvas.drawText(text_object)
                
            # Define as configurações de texto
            text_object = pdf_canvas.beginText(x=100, y=altura_linha, )
            altura_linha = (altura_linha - 35)
            text_object.setFont("Fonte", 10)
            text_object.setFillColor(colors.black)
            # Adiciona o texto personalizável
            text_object.textLine(' ')
            # Desenha o texto no canvas
            pdf_canvas.drawText(text_object)
            
    # Fecha o arquivo PDF
    pdf_canvas.save()
    
    
def analisa_subpastas(caminho_subpasta, nome_subpasta, subpasta_arquivos):
    # para cada arquivo na subpasta
    for arquivo in os.listdir(caminho_subpasta):
        if arquivo.endswith('.pdf'):
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
                          f'Arquivo "{arquivo}" protegido encontrado em: {os.path.join(caminho_subpasta, arquivo)}'
                          f'Para que o processo automatizado mescle PDF protegido, a senha deve ser informada no nome do arquivo, por exemplo:\n'
                          f'Senha 1234567.pdf\n')
                    break
            
            # se a subpasta já consta no dicionário, adiciona mais um pdf
            if nome_subpasta in subpasta_arquivos:
                subpasta_arquivos[nome_subpasta].append(arquivo)
            # se não cria uma nova subpasta dentro do dicionário já adicionando o pdf
            else:
                subpasta_arquivos[nome_subpasta] = [arquivo]
    
    return True, subpasta_arquivos


def analisa_documentos(pasta_inicial):
    arquivos = []
    # para cada arquivo na subpasta
    for arquivo in os.listdir(pasta_inicial):
        # print(arquivo)
        if arquivo.endswith('.pdf'):
            # verifica se o arquivo tem senha, se tiver tira a senha dele
            try:
                with fitz.open(os.path.join(pasta_inicial, arquivo)) as pdf:
                    arquivo_sem_senha = pdf
            except:
                try:
                    password = re.compile(r'Senha.(\w+)').search(arquivo).group(1)
                    decryption(os.path.join(caminho_subpasta, arquivo), os.path.join(caminho_subpasta, arquivo), password)
                except:
                    # se tiver senha, mas a senha não for informada no nome do arquivo, pula para próxima subpasta
                    alert(f'Não é possível criar E-Book de {nome_subpasta}.\n'
                          f'Existe PDF protegido sem a senha informada no nome do arquivo.\n'
                          f'Arquivo "{arquivo}" protegido encontrado em: {os.path.join(caminho_subpasta, arquivo)}'
                          f'Para que o processo automatizado mescle PDF protegido, a senha deve ser informada no nome do arquivo, por exemplo:\n'
                          f'Senha 1234567.pdf\n')
                    break
            
            arquivos.append(arquivo)

    return False, arquivos


def cria_ebook(window, subpasta, nomes_arquivos, pasta_final):
    window['-Mensagens-'].update(f'Criando arquivos...')
    achou = 'não'
    for arquivo in nomes_arquivos:
        if re.compile(r'imagem-recibo').search(arquivo):
            achou = 'sim'
    
    if achou == 'não':
        if not subpasta:
            alert(text=f'Não foi encontrado recibo de entrega DIRPF')
            return False
        else:
            alert(text=f'Não foi encontrado recibo de entrega DIRPF na pasta {subpasta}')
            return False
        
    contador = 4
    lista_arquivos = []
    infos_resumo = []
    # cria uma cópia numerada dos PDFs
    for arquivo in nomes_arquivos:
        if not subpasta:
            abre_pdf = os.path.join(pasta_inicial, arquivo)
        else:
            abre_pdf = os.path.join(pasta_inicial, subpasta, arquivo)
    
        if re.compile(r'imagem-recibo').search(arquivo):
            # busca o nome do declarante
            with fitz.open(abre_pdf) as pdf:
                conteudo_recibo = ''
                for page in pdf:
                    texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                    conteudo_recibo += texto_pagina
                
                # print(conteudo_recibo)
                # time.sleep(30)
                
                try:
                    titulo_resumo = re.compile('(DECLARAÇÃO DE AJUSTE ANUAL - .+)').search(conteudo_recibo).group(1)
                    infos_resumo.append((0, f'RESUMO DA {titulo_resumo.replace(" - ", " ")}'))
                    infos_resumo.append((1, ''))
                    
                    total_rendimentos_tributaveis = re.compile('(.+,\d+)\nTOTAL RENDIMENTOS TRIBUTÁVEIS').search(conteudo_recibo).group(1)
                    infos_resumo.append((2, f'Total dos rendimentos tributáveis no ano: R${total_rendimentos_tributaveis}'))
                    
                    imposto_devido = re.compile('(.+,\d+)\nIMPOSTO DEVIDO').search(conteudo_recibo).group(1)
                    infos_resumo.append((3, f'Imposto devido no ano: R${imposto_devido}'))
                    
                    nome_declarante = re.compile(r'Nome do declarante\n.+\n(.+)').search(conteudo_recibo).group(1)
                    
                    pagar_restituir = re.compile(r'(.+,\d+)\nIMPOSTO A RESTITUIR').search(conteudo_recibo).group(1)
                    if pagar_restituir == '0,00':
                        pagar_restituir = re.compile(r'IMPOSTO A PAGAR(\n.+){2}\n(.+,\d\d)').search(conteudo_recibo).group(2)
                        if pagar_restituir == '0,00':
                            pagar_restituir = ''
                        else:
                            pagar_restituir = f'Saldo a pagar: R$ {pagar_restituir}'
                            infos_resumo.append((4, f'Imposto a restituir: R$ 0,00'))
                            infos_resumo.append((5, pagar_restituir.replace('Saldo a pagar', 'Imposto a pagar')))
                    else:
                        pagar_restituir = f'Saldo a restituir: R$ {pagar_restituir}'
                        infos_resumo.append((4, pagar_restituir.replace('Saldo a restituir', 'Imposto a restituir')))
                        infos_resumo.append((5, f'Imposto a pagar: R$ 0,00'))
                except:
                    alert(text=f'Não foi possível encontrar o nome do declarante no recibo de entrega do IRPF:\n'
                               f'{abre_pdf}\n'
                               f'Verifique a forma que o PDF foi salvo e tente novamente.')
                    return False
                    
            # cria a capa
            if not subpasta:
                nome_do_arquivo = os.path.join(pasta_inicial, f'Capa E-book {nome_declarante}.pdf')
            else:
                nome_do_arquivo = os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf')
            
            caminho_da_imagem = "Assets\DIRPF_capa.png"
            create_pdf(nome_do_arquivo, caminho_da_imagem, nome_declarante, pagar_restituir)
            # adiciona o arquivo da capa na lista da subpasta
            nomes_arquivos.append(f'Capa E-book {nome_declarante}.pdf')
            
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
                
            lista_arquivos.append(0)
            
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
                
            lista_arquivos.append(2)
            continue
        
        if re.compile(r'imagem-declaracao').search(arquivo):
            # busca o nome do declarante
            with fitz.open(abre_pdf) as pdf:
                conteudo_declaracao = ''
                for page in pdf:
                    texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                    conteudo_declaracao += texto_pagina
                
                # print(conteudo_declaracao)
                # time.sleep(30)

                aliquota_efetiva = re.compile('(.+,\d+)\n.+\nAliquota efetiva \(%\)\nBase de cálculo do imposto').search(conteudo_declaracao)
                if not aliquota_efetiva:
                    aliquota_efetiva = re.compile('(.+,\d+)\nAliquota efetiva \(%\)\nTipo de Conta').search(conteudo_declaracao)
                    if not aliquota_efetiva:
                        try:
                            alert(text=f'Não foi encontrada a Aliquota efetiva (%) na declaração do declarante: {nome_declarante}.\n'
                                        'Provável erro de layout da página, a informação não será incluída no resumo localizado na página 1')
                        except:
                            alert(text=f'Não foi possível encontrar o nome do declarante no recibo de entrega do IRPF:\n'
                                       f'Verifique a forma que o PDF foi salvo e tente novamente.')
                            return False
                    else:
                        aliquota_efetiva = aliquota_efetiva.group(1)
                        infos_resumo.append((6, f'Aliquota efetiva no ano em percentual (%): {aliquota_efetiva}'))
                else:
                    aliquota_efetiva = aliquota_efetiva.group(1)
                    infos_resumo.append((6, f'Aliquota efetiva no ano em percentual (%): {aliquota_efetiva}'))
                    
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
                
            lista_arquivos.append(3)
            continue
        
        if re.compile(r'INFORME').search(arquivo.upper()):

            with fitz.open(abre_pdf) as pdf:
                texto_arquivo = ''
                for page in pdf:
                    texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                    texto_arquivo += texto_pagina
                
                informe_empresa = re.compile(r'COMPROVANTE DE RENDIMENTOS PAGOS E DE IMPOSTO SOBRE A RENDA RETIDO NA FONTE').search(texto_arquivo.upper())
                if not informe_empresa:
                    informe_empresa = re.compile(r'Comprovante de Rendimentos Pagos e de\nImposto sobre a Renda Retido na Fonte').search(texto_arquivo)
                
                if informe_empresa:
                    if not subpasta:
                        shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                    else:
                        shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        
                    lista_arquivos.append(contador)
                    contador += 1
    
    cria_pagina_resumo(infos_resumo, 'Arquivos para mesclar\\1.pdf')
    lista_arquivos.append(1)
    
    for arquivo in nomes_arquivos:
        if not subpasta:
            abre_pdf = os.path.join(pasta_inicial, arquivo)
        else:
            abre_pdf = os.path.join(pasta_inicial, subpasta, arquivo)
            
        if re.compile(r'INFORME').search(arquivo.upper()):
            with fitz.open(abre_pdf) as pdf:
                texto_arquivo = ''
                for page in pdf:
                    texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                    texto_arquivo += texto_pagina
                
                informe_empresa = re.compile(r'COMPROVANTE DE RENDIMENTOS PAGOS E DE IMPOSTO SOBRE A RENDA RETIDO NA FONTE').search(texto_arquivo.upper())
                if not informe_empresa:
                    informe_empresa = re.compile(r'Comprovante de Rendimentos Pagos e de\nImposto sobre a Renda Retido na Fonte').search(texto_arquivo)
                
                if not informe_empresa:
                    if not subpasta:
                        shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                    else:
                        shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        
                    lista_arquivos.append(contador)
                    contador += 1
    
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
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                
            lista_arquivos.append(contador)
            contador += 1
        
    # ordena a lista de arquivos
    lista_arquivos = sorted(lista_arquivos, key=chave_numerica)
    
    # mescla os arquivos
    pdf_merger = PyPDF2.PdfMerger()
    for count, arquivo in enumerate(lista_arquivos, start=1):
        caminho_completo = os.path.join('Arquivos para mesclar', f'{arquivo}.pdf')
        pdf_merger.append(caminho_completo)
        
        if not subpasta:
            window['-progressbar-'].update_bar(count, max=int(len(lista_arquivos)))
            window['-Progresso_texto-'].update(str(round(float(count) / int(len(lista_arquivos)) * 100, 1)) + '%')
            window.refresh()
        
    pdf_merger.append(os.path.join('Assets', 'AGRADECIMENTO POR ESCOLHER A VEIGA & POSTAL - IPRF TIMBRADO.pdf'))
    # cria o e-book
    unificado_pdf = os.path.join(pasta_final, f'E-BOOK DIRPF - {nome_declarante}.pdf')
    while True:
        try:
            pdf_merger.write(unificado_pdf)
            pdf_merger.close()
            break
        except:
            alert('Atualização de e-book falhou.\nCaso exista algum e-book aberto, por gentileza feche para que ele seja atualizado.')
            
    window['-Mensagens-'].update(f'Finalizando, aguarde...')
    coloca_marca_dagua(unificado_pdf)
    return True
    

def coloca_marca_dagua(unificado_pdf):
    # Abre o arquivo de entrada PDF
    with open(unificado_pdf, 'rb') as input_file:
        input_pdf_reader = PyPDF2.PdfReader(input_file)
        output_pdf_writer = PyPDF2.PdfWriter()
        
        # Carrega a imagem da marca d'água
        watermark = ImageReader('Assets/Logo_VP.png')
        
        # Loop através de cada página do PDF de entrada
        for page_number in range(len(input_pdf_reader.pages)):
            # Ignora a primeira e a última página
            if page_number == 0 or page_number == len(input_pdf_reader.pages) - 1:
                input_page = input_pdf_reader.pages[page_number]
                output_pdf_writer.add_page(input_page)
                continue
            
            input_page = input_pdf_reader.pages[page_number]
            width = input_page.mediabox.upper_right[0] - input_page.mediabox.lower_left[0]
            height = input_page.mediabox.upper_right[1] - input_page.mediabox.lower_left[1]
            
            tamanho = (float(width), float(height))
            # Adiciona a marca d'água em todas as outras páginas
            packet = io.BytesIO()
            # print(tamanho)
            can = canvas.Canvas(packet, pagesize=tamanho)
            can.drawImage(watermark, tamanho[0] - 130, 30, width=95, height=25, mask='auto')
            can.save()
            
            packet.seek(0)
            overlay = PyPDF2.PdfReader(packet)
            input_page.merge_page(overlay.pages[0])
            output_pdf_writer.add_page(input_page)
        
        # Salva o PDF resultante
        with open(unificado_pdf, 'wb') as output_file:
            output_pdf_writer.write(output_file)


def run(window, pasta_inicial, pasta_final):
    # inicia o dicionário
    arquivos = None
    nomes_arquivos = None
    subpasta_arquivos = {}
    nome_subpasta = True
    # itera sobre todas as subpastas dentro da pasta mestre
    for count, nome_subpasta in enumerate(os.listdir(pasta_inicial), start=1):
        window['-Mensagens-'].update(f'Analisando arquivos...')
        caminho_subpasta = os.path.join(pasta_inicial, nome_subpasta)
        
        # Verifica se é uma pasta
        if os.path.isdir(caminho_subpasta):
            print(f"Entrando na subpasta: {nome_subpasta}")
            nome_subpasta, subpasta_arquivos = analisa_subpastas(caminho_subpasta, nome_subpasta, subpasta_arquivos)
        
        else:
            if caminho_subpasta.endswith('.pdf'):
                nome_subpasta, arquivos = analisa_documentos(pasta_inicial)
            else:
                continue
            break
            
        window['-progressbar-'].update_bar(count, max=int(len(os.listdir(pasta_inicial))))
        window['-Progresso_texto-'].update(str(round(float(count) / int(len(os.listdir(pasta_inicial))) * 100, 1)) + '%')
        window.refresh()
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
     
    # ---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    os.makedirs('Arquivos para mesclar', exist_ok=True)
    # limpa a pasta de cópias de arquivos
    resultado = False
    for arquivo in os.listdir('Arquivos para mesclar'):
        os.remove(os.path.join('Arquivos para mesclar', arquivo))
    
    window['-progressbar-'].update_bar(0)
    window['-Progresso_texto-'].update('')
    
    print(nome_subpasta)
    if not nome_subpasta:
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
        
        resultado = cria_ebook(window, False, arquivos, pasta_final)
        
        # limpa a pasta de cópias de arquivos
        for arquivo in os.listdir('Arquivos para mesclar'):
            os.remove(os.path.join('Arquivos para mesclar', arquivo))
        
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
        
    else:
        # para cada subpasta
        for count, (subpasta, nomes_arquivos) in enumerate(subpasta_arquivos.items(), start=1):
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                print('ENCERRAR')
                return
            
            resultado = cria_ebook(window, subpasta, nomes_arquivos, pasta_final)
            
            window['-progressbar-'].update_bar(count, max=int(len(subpasta_arquivos.items())))
            window['-Progresso_texto-'].update(str(round(float(count) / int(len(subpasta_arquivos.items())) * 100, 1)) + '%')
            window.refresh()
            
            # limpa a pasta de cópias de arquivos
            for arquivo in os.listdir('Arquivos para mesclar'):
                os.remove(os.path.join('Arquivos para mesclar', arquivo))
            
            if event == '-encerrar-' or event == sg.WIN_CLOSED:
                print('ENCERRAR')
                return
    
    if resultado:
        alert(text='PDFs unificados com sucesso.')
    
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    layout = [
        [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True)],
        [sg.Text('')],
        [sg.Text('Selecione a pasta que contenha os arquivos PDF para analisar:')],
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
        
        if len(os.listdir(pasta_inicial)) < 1:
            alert(text=f'Nenhum arquivo PDF encontrado na pasta selecionada.')
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
        
        try:
            # Chama a função que executa o script
            run(window, pasta_inicial, pasta_final)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            escreve_doc(erro)
            window['Log do sistema'].update(disabled=False)
            alert(text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
        
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
            os.startfile('Log')
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Cria E-book DIRPF.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
        
        elif event == '-abrir_resultados-':
            os.startfile(pasta_final)
    
    window.close()
    
    
    