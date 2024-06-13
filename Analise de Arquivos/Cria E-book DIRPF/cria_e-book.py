# -*- coding: utf-8 -*-
import json, time, shutil, io, fitz, PyPDF2, re, os, sys, traceback, PySimpleGUI as sg
from PIL import Image
from threading import Thread
from pyautogui import alert, confirm
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from PIL import Image, ImageDraw
from PySimpleGUI import BUTTON_TYPE_READ_FORM, FILE_TYPES_ALL_FILES, theme_background_color, theme_button_color, BUTTON_TYPE_BROWSE_FILE, BUTTON_TYPE_BROWSE_FOLDER
from base64 import b64encode

# Carregar a fonte TrueType (substitua 'sua_fonte.ttf' pelo caminho da sua fonte)
pdfmetrics.registerFont(TTFont('Fonte', 'Assets\HankenGrotesk-SemiBold.ttf'))
tamanho_da_pagina = (672,950)

# Define o ícone global da aplicação
sg.set_global_icon('Assets/EB_icon.ico')
dados_modo = 'Assets/modo.txt'
dados_elementos = 'Log/window_values.json'
controle_botoes = 'Log/Buttons.txt'

versao_atual = '2.0'
caminho_arquivo = 'T:\ROBÔ\_Executáveis\Cria E-book\Cria E-book DIRPF.exe'

cr = open('T:\\ROBÔ\\_Executáveis\\Cria E-book\\Versão.txt', 'r', encoding='utf-8').read()
if str(cr) != versao_atual:
    while True:
        atualizar = confirm(text='Existe uma nova versão do programa, deseja atualizar agora?', buttons=('Novidades da atualização', 'Sim', 'Não'))
        if atualizar == 'Novidades da atualização':
            os.startfile('T:\\ROBÔ\\_Executáveis\\Cria E-book\\Sobre.pdf')
        if atualizar == 'Sim':
            break
        elif atualizar == 'Não':
            break
    
    if atualizar == 'Sim':
        os.startfile(caminho_arquivo)
        sys.exit()
        

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
    reader = PyPDF2.PdfReader(pdfFile, strict=False)
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


def cria_capa(output_path, image_path, text, text_2):
    text = quebra_de_linha(text, caracteres_por_linha=16)
    
    # Abre o arquivo PDF para gravação
    pdf_canvas = canvas.Canvas(output_path, pagesize=tamanho_da_pagina)

    # Adiciona a imagem à posição desejada
    pdf_canvas.drawInlineImage(image_path, 0, 0, width=tamanho_da_pagina[0], height=tamanho_da_pagina[1])
    
    def coloca_nome_e_imposto(pdf_canvas, text, text_2):
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
        
        return pdf_canvas
        
    pdf_canvas = coloca_nome_e_imposto(pdf_canvas, text, text_2)
    
    # Fecha o arquivo PDF
    pdf_canvas.save()
    

def cria_pagina_resumo(infos_resumo, output_path):
    titulo_equipe = "Assets\\Titulo Equipe.txt"
    f = open(titulo_equipe, 'r', encoding='utf-8')
    titulos_equipe = f.readline().split('/')
    
    equipe = "Assets\\Equipe.txt"
    f = open(equipe, 'r', encoding='utf-8')
    equipe = f.readline()
    equipes = equipe.split('|')
    
    # c.rect(x, y, width, height, fill=1)  #draw rectangle
    infos_resumo = sorted(infos_resumo)
    data_list = []
    title_data = [['SEM TITULO']]
    for info in infos_resumo:
        if info[0] == 0:
            text = info[1].upper().split('-')
            title_data = [[text[0]], [text[1]]]
            print(text)
        else:
            print(info)
            data_list.append((info[1].upper(), info[2]))
    
    # Criação do documento PDF
    pdf_right_aligned = SimpleDocTemplate(output_path, pagesize=tamanho_da_pagina, topMargin=44, bottonMargin=0)
    tabela = []
    
    
    def cria_tabela_resumo(tabela, title_data, data_list):
        # Estilo para o título da tabela
        title_style_left_aligned = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Fonte'),
            ('FONTSIZE', (0, 0), (-1, -1), 15),  # Mesmo tamanho de fonte para ambas as linhas
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
        ])
        
        title_table_left_aligned = Table(title_data, colWidths=[580])  # Ajuste da largura para alinhamento à esquerda
        title_table_left_aligned.setStyle(title_style_left_aligned)
        tabela.append(title_table_left_aligned)
        
        # Adicionar espaço
        tabela.append(Table([['']], colWidths=[None], rowHeights=[20]))
        
        # Estilo para a tabela com valores alinhados à direita
        right_align_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.white),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (-1, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (-1, -1), 'Fonte'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 20),
            ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.black),  # Linha fina abaixo de cada item
        ])
        
        # Definindo a tabela com alinhamento à direita
        right_aligned_table = Table(data_list, colWidths=[330, 250])
        right_aligned_table.setStyle(right_align_style)
        
        # Adicionar a tabela aos elementos do documento
        tabela.append(right_aligned_table)
        
        return tabela
    
    
    def coloca_equipe(tabela):
        listas = []
        for count, titulo in enumerate(titulos_equipe):
            listas.append((titulo, equipes[count]))
        
        altura_linha = (950 - ((len(title_data) * 20) + ((len(data_list)+1) * 30)) - (((len(equipes)+1) * 12) + (len(titulo_equipe) * 16) + (35*4)))
        # Adicionar espaço
        tabela.append(Table([['']], colWidths=[None], rowHeights=[altura_linha]))
        
        for lista in listas:
            _titulo = lista[0]
            _equipe = lista[1].split('/')
            data_equipe = []
            for info in _equipe:
                data_equipe.append([info])

            # Estilo para o título da tabela
            title_style_left_aligned = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Fonte'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),  # Mesmo tamanho de fonte para ambas as linhas
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
            ])
            # Estilo para a tabela com valores alinhados à direita
            right_align_style = TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Fonte'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),  # Mesmo tamanho de fonte para ambas as linhas
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ])
            
            # Adicionar espaço
            tabela.append(Table([['']], colWidths=[None], rowHeights=[22]))
            # TITULO
            title_table_left_aligned = Table([[_titulo]], colWidths=[580])  # Ajuste da largura para alinhamento à esquerda
            title_table_left_aligned.setStyle(title_style_left_aligned)
            tabela.append(title_table_left_aligned)
            # Adicionar espaço
            tabela.append(Table([['']], colWidths=[None], rowHeights=[13]))
            # LISTA
            right_aligned_table = Table(data_equipe, colWidths=[580])
            right_aligned_table.setStyle(right_align_style)
            tabela.append(right_aligned_table)
        
        return tabela
    
    
    tabela = cria_tabela_resumo(tabela, title_data, data_list)
    tabela = coloca_equipe(tabela)
    
    # Construir o PDF com os elementos
    pdf_right_aligned.build(tabela)

 
def analisa_subpastas(caminho_subpasta, nome_subpasta, subpasta_arquivos):
    # para cada arquivo na subpasta
    for arquivo in os.listdir(caminho_subpasta):
        if arquivo.endswith('.pdf'):
            print(arquivo)
            # verifica se o arquivo tem senha, se tiver tira a senha dele
            try:
                with fitz.open(os.path.join(caminho_subpasta, arquivo)) as pdf:
                    for page in pdf:
                        break
            except:
                try:
                    password = re.compile(r'SENHA.(\w+)').search(arquivo.upper()).group(1)
                    decryption(os.path.join(caminho_subpasta, arquivo), os.path.join(caminho_subpasta, arquivo), password)
                except:
                    # se tiver senha, mas a senha não for informada no nome do arquivo, pula para próxima subpasta
                    alert(f'Não é possível criar E-Book de {nome_subpasta}.\n\n'
                          f'Existe PDF protegido sem a senha informada no nome do arquivo.\n\n'
                          f'Arquivo "{arquivo}" protegido encontrado em: {os.path.join(caminho_subpasta, arquivo)}\n\n'
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
        if arquivo.endswith('.pdf'):
            # verifica se o arquivo tem senha, se tiver tira a senha dele
            try:
                with fitz.open(os.path.join(pasta_inicial, arquivo)) as pdf:
                    for page in pdf:
                        break
            except:
                try:
                    print(arquivo)
                    password = re.compile(r'SENHA.(\w+)').search(arquivo.upper()).group(1)
                    decryption(os.path.join(pasta_inicial, arquivo), os.path.join(pasta_inicial, arquivo), password)
                except:
                    # se tiver senha, mas a senha não for informada no nome do arquivo, pula para próxima subpasta
                    alert(f'Não é possível criar E-Book de {pasta_inicial}.\n'
                          f'Existe PDF protegido sem a senha informada no nome do arquivo.\n'
                          f'Arquivo "{arquivo}" protegido encontrado em: {os.path.join(pasta_inicial, arquivo)}'
                          f'Para que o processo automatizado mescle PDF protegido, a senha deve ser informada no nome do arquivo, por exemplo:\n'
                          f'Senha 1234567.pdf\n')
                    break
            
            arquivos.append(arquivo)

    return False, arquivos


def remove_metadata(input_pdf, output_pdf):
    with open(input_pdf, 'rb') as input_file:
        pdf_reader = PyPDF2.PdfReader(input_file, strict=False)
        pdf_writer = PyPDF2.PdfWriter()
        
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)
            
        # Add the metadata
        pdf_writer.add_metadata(
            {
                "/Author": "Dev",
                "/Producer": "Cria_ebook",
            }
        )
        
    # Cria um novo arquivo PDF sem metadados
    with open(output_pdf, 'wb') as output_file:
        pdf_writer.write(output_file)
        

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

    # lista para os arquivos que serão mesclados
    lista_arquivos = []
    # lista para a página de resumo
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
                    # captura o titulo para a página de resumo
                    titulo_resumo = re.compile('(DECLARAÇÃO DE AJUSTE ANUAL - .+)').search(conteudo_recibo).group(1)
                    infos_resumo.append((0, f'RESUMO DA {titulo_resumo.replace(" - ", "-")}'))
                    
                    # adiciona na lista da página de resumo o total dos rendimentos tributáveis
                    total_rendimentos_tributaveis = re.compile('(.+,\d+)\nTOTAL RENDIMENTOS TRIBUTÁVEIS').search(conteudo_recibo).group(1)
                    infos_resumo.append((2, 'Total dos rendimentos tributáveis no ano', f'R$ {total_rendimentos_tributaveis}'))
                    # adiciona na lista da página de resumo o imposto devido
                    imposto_devido = re.compile('(.+,\d+)\nIMPOSTO DEVIDO').search(conteudo_recibo).group(1)
                    infos_resumo.append((3, 'Imposto devido no ano', f'R$ {imposto_devido}'))
                    
                    # captura o nome do declarante
                    nome_declarante = re.compile(r'Nome do declarante\n.+\n(.+)').search(conteudo_recibo).group(1)
                    print(f'\n{nome_declarante}')
                    
                    # adiciona na lista da página de resumo o imposto a restituir e o imposto a pagar
                    # se tiver a restituir adiciona na variável 'pagar_restituir' para escrever na capa com o nome
                    # se tiver a pagar adiciona na variável 'pagar_restituir' para escrever na capa com o nome
                    # a variável 'pagar_restituir' terá apenas uma das duas informações
                    pagar_restituir = re.compile(r'(.+,\d+)\nIMPOSTO A RESTITUIR').search(conteudo_recibo).group(1)
                    if pagar_restituir == '0,00':
                        pagar_restituir = re.compile(r'IMPOSTO A PAGAR(\n.+){2}\n(.+,\d\d)').search(conteudo_recibo).group(2)
                        if pagar_restituir == '0,00':
                            pagar_restituir = ''
                            infos_resumo.append((4, 'Imposto a restituir', 'R$ 0,00'))
                            infos_resumo.append((5, 'Imposto a pagar', 'R$ 0,00'))
                        else:
                            infos_resumo.append((4, 'Imposto a restituir', 'R$ 0,00'))
                            infos_resumo.append((5, 'Imposto a pagar', f'R$ {pagar_restituir}'))
                            pagar_restituir = f'Saldo a pagar: R$ {pagar_restituir}'
                    else:
                        infos_resumo.append((4, 'Imposto a restituir', f'R$ {pagar_restituir}'))
                        infos_resumo.append((5, 'Imposto a pagar', 'R$ 0,00'))
                        pagar_restituir = f'Saldo a restituir: R$ {pagar_restituir}'
                except:
                    alert(text=f'Não foi possível encontrar o nome do declarante no recibo de entrega do IRPF:\n\n'
                               f'{abre_pdf}\n\n'
                               f'Verifique a forma que o PDF foi salvo e tente novamente.')
                    return False
                    
            # cria a capa
            if not subpasta:
                nome_do_arquivo = os.path.join(pasta_inicial, f'Capa E-book {nome_declarante}.pdf')
            else:
                nome_do_arquivo = os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf')
            
            caminho_da_imagem = "Assets\DIRPF_capa.png"
            cria_capa(nome_do_arquivo, caminho_da_imagem, nome_declarante, pagar_restituir)
            # adiciona o arquivo da capa na lista da subpasta
            nomes_arquivos.append(f'Capa E-book {nome_declarante}.pdf')
            
            # cria uma cópia númerada da capa
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, subpasta, f'Capa E-book {nome_declarante}.pdf'), os.path.join('Arquivos para mesclar', '0.pdf'))
            
            # adiciona a capa na lista de arquivos para mesclar
            lista_arquivos.append(0)
            
            # cria uma cópia númerada do recibo
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '2.pdf'))
            # adiciona o recibo na lista de arquivos para mesclar
            lista_arquivos.append(2)
            continue
        
        if re.compile(r'imagem-declaracao').search(arquivo):
            with fitz.open(abre_pdf) as pdf:
                conteudo_declaracao = ''
                for page in pdf:
                    texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                    conteudo_declaracao += texto_pagina
                
                # print(conteudo_declaracao)
                # time.sleep(30)
                
                # adiciona na lista da página de resumo a aliquota efetiva
                # existem duas variantes da informação no PDF, dependendo do tipo de declaração escolhida, OPÇÃO PELAS DEDUÇÕES LEGAIS ou OPÇÃO PELO DESCONTO SIMPLIFICADO
                regex_aliquotas = [r'(.+,\d+)\n.+\nAliquota efetiva \(%\)\nBase de cálculo do imposto',
                                   r'(.+,\d+)\nAliquota efetiva \(%\)\nTipo de Conta',
                                   r'(.+,\d+)\nAliquota efetiva \(%\)\nPágina \d']
                for regex_a in regex_aliquotas:
                    aliquota_efetiva = re.compile(regex_a).search(conteudo_declaracao)
                    if aliquota_efetiva:
                        aliquota_efetiva = aliquota_efetiva.group(1)
                        infos_resumo.append((6, f'Aliquota efetiva no ano em percentual (%)', aliquota_efetiva))
                        break
                if not aliquota_efetiva:
                    alert(text=f'Não foi encontrada a Aliquota efetiva (%) no arquivo: {arquivo}.\n\n'
                               'Provável erro de layout da página.\n\n'
                               'A informação não será inserida na página de resumo, verifique a forma que o PDF foi salvo e tente novamente.')
            
            # cria uma cópia númerada da declaração
            if not subpasta:
                shutil.copy(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', '3.pdf'))
            # adiciona a declaração na lista de arquivos para mesclar
            lista_arquivos.append(3)
            continue
        
        # Adiciona os Informes de rendimento da empresa depois da declaração
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
                        # remove_metadata(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                    else:
                        shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        # remove_metadata(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        
                    lista_arquivos.append(contador)
                    contador += 1
    
    # cria a página de resumo
    cria_pagina_resumo(infos_resumo, 'Arquivos para mesclar\\1.pdf')
    lista_arquivos.append(1)
    
    # percorre os arquivos novamente, mas dessa vês adiciona informes de banco e outros
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
                        # remove_metadata(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                    else:
                        shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        # remove_metadata(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                        
                    lista_arquivos.append(contador)
                    contador += 1
    
    # percorre os arquivos novamente, adicionando os outros documentos
    for arquivo in nomes_arquivos:
        if not subpasta:
            abre_pdf = os.path.join(pasta_inicial, arquivo)
        else:
            abre_pdf = os.path.join(pasta_inicial, subpasta, arquivo)
            
        print(f'\n{arquivo}')
        with open(abre_pdf, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file, strict=False)
            # Verifica se o PDF possui metadados
            meta = pdf_reader.metadata
        
        if meta:
            for value in meta.items():
                print(value)
            
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
                # remove_metadata(os.path.join(pasta_inicial, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
            else:
                shutil.copy(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
                # remove_metadata(os.path.join(pasta_inicial, subpasta, arquivo), os.path.join('Arquivos para mesclar', str(contador) + '.pdf'))
            
            lista_arquivos.append(contador)
            contador += 1
        
    # ordena a lista de arquivos
    lista_arquivos = sorted(lista_arquivos, key=chave_numerica)
    
    # mescla os arquivos
    pdf_merger = PyPDF2.PdfMerger(strict=False)
    
    for count, arquivo in enumerate(lista_arquivos, start=1):
        caminho_completo = os.path.join('Arquivos para mesclar', f'{arquivo}.pdf')
        pdf_merger.append(caminho_completo)
        
        if not subpasta:
            window['-progressbar-'].update_bar(count, max=int(len(lista_arquivos)))
            window['-Progresso_texto-'].update(str(round(float(count) / int(len(lista_arquivos)) * 100, 1)) + '%')
            window.refresh()
    
    pdf_merger.append(os.path.join('Assets', 'AGRADECIMENTO - IPRF TIMBRADO.pdf'))

    # cria o e-book
    unificado_pdf = os.path.join(pasta_final, f'E-BOOK DIRPF - {nome_declarante}.pdf')
    while True:
        try:
            pdf_merger.write(unificado_pdf)
            pdf_merger.close()
            break
        except:
            pdf_merger.close()
            alert('Atualização de e-book falhou.\n\nCaso exista algum e-book aberto, por gentileza feche para que ele seja atualizado.')
            return False
            
    window['-Mensagens-'].update(f'Finalizando, aguarde...')
    coloca_marca_dagua(unificado_pdf)
    return True
    

def coloca_marca_dagua(unificado_pdf):
    # Abre o arquivo de entrada PDF
    with open(unificado_pdf, 'rb') as input_file:
        input_pdf_reader = PyPDF2.PdfReader(input_file, strict=False)
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
            overlay = PyPDF2.PdfReader(packet, strict=False)
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
        print('Limpando pasta de apoio')
        os.remove(os.path.join('Arquivos para mesclar', arquivo))
    
    window['-progressbar-'].update_bar(0)
    window['-Progresso_texto-'].update('')
    
    print(nome_subpasta)
    if not nome_subpasta:
        if event == '-encerrar-' or event == sg.WIN_CLOSED:
            print('ENCERRAR')
            return
        
        resultado = cria_ebook(window, False, arquivos, pasta_final)
        if not resultado:
            return
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
            if not resultado:
                return
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


# Função para salvar os valores dos elementos em um arquivo JSON
def save_values(values, filename=dados_elementos):
    with open(filename, 'w', encoding='latin-1') as f:
        json.dump(values, f)


# Função para carregar os valores dos elementos de um arquivo JSON
def load_values(filename=dados_elementos):
    if os.path.exists(filename):
        with open(filename, 'r', encoding='latin-1') as f:
            return json.load(f)
    return {}


if __name__ == '__main__':
    def RoundedButton(button_text=' ', corner_radius=0.1, button_type=BUTTON_TYPE_READ_FORM, target=(None, None), tooltip=None, file_types=FILE_TYPES_ALL_FILES, initial_folder=None, default_extension='',
                      disabled=False, change_submits=False, enable_events=False, image_size=(None, None), image_subsample=None, border_width=None, size=(None, None), auto_size_button=None,
                      button_color=None, disabled_button_color=None, highlight_colors=None, mouseover_colors=(None, None), use_ttk_buttons=None, font=None, bind_return_key=False, focus=False,
                      pad=None, key=None, right_click_menu=None, expand_x=False, expand_y=False, visible=True, metadata=None):
        if None in size:
            multi = 5
            size = (((len(button_text) if size[0] is None else size[0]) * 5 + 20) * multi,
                    20 * multi if size[1] is None else size[1])
        if button_color is None:
            button_color = theme_button_color()
        btn_img = Image.new('RGBA', size, (0, 0, 0, 0))
        corner_radius = int(corner_radius / 2 * min(size))
        poly_coords = (
            (corner_radius, 0),
            (size[0] - corner_radius, 0),
            (size[0], corner_radius),
            (size[0], size[1] - corner_radius),
            (size[0] - corner_radius, size[1]),
            (corner_radius, size[1]),
            (0, size[1] - corner_radius),
            (0, corner_radius),
        )
        pie_coords = [
            [(size[0] - corner_radius * 2, size[1] - corner_radius * 2, size[0], size[1]),
             [0, 90]],
            [(0, size[1] - corner_radius * 2, corner_radius * 2, size[1]), [90, 180]],
            [(0, 0, corner_radius * 2, corner_radius * 2), [180, 270]],
            [(size[0] - corner_radius * 2, 0, size[0], corner_radius * 2), [270, 360]],
        ]
        brush = ImageDraw.Draw(btn_img)
        brush.polygon(poly_coords, button_color[1])
        for coord in pie_coords:
            brush.pieslice(coord[0], coord[1][0], coord[1][1], button_color[1])
        data = io.BytesIO()
        btn_img.thumbnail((size[0] // 3, size[1] // 3), resample=Image.LANCZOS)
        btn_img.save(data, format='png', quality=95)
        btn_img = b64encode(data.getvalue())
        return sg.Button(button_text=button_text, button_type=button_type, target=target, tooltip=tooltip,
                         file_types=file_types, initial_folder=initial_folder, default_extension=default_extension,
                         disabled=disabled, change_submits=change_submits, enable_events=enable_events,
                         image_data=btn_img, image_size=image_size,
                         image_subsample=image_subsample, border_width=border_width, size=size,
                         auto_size_button=auto_size_button, button_color=(button_color[0], theme_background_color()),
                         disabled_button_color=disabled_button_color, highlight_colors=highlight_colors,
                         mouseover_colors=mouseover_colors, use_ttk_buttons=use_ttk_buttons, font=font,
                         bind_return_key=bind_return_key, focus=focus, pad=pad, key=key, right_click_menu=right_click_menu,
                         expand_x=expand_x, expand_y=expand_y, visible=visible, metadata=metadata)
        
    # limpa o arquivo que salva o estado dos elementos preenchidos e o estado anterior do botão de abrir os resultados
    with open(dados_elementos, 'w', encoding='utf-8') as f:
        f.write('')
    with open(controle_botoes, 'w', encoding='utf-8') as f:
        f.write('')
        
    sg.LOOK_AND_FEEL_TABLE['tema_claro'] = {'BACKGROUND': '#F8F8F8',
                                            'TEXT': '#000000',
                                            'INPUT': '#F8F8F8',
                                            'TEXT_INPUT': '#000000',
                                            'SCROLL': '#F8F8F8',
                                            'BUTTON': ('#000000', '#F8F8F8'),
                                            'PROGRESS': ('#fca400', '#D7D7D7'),
                                            'BORDER': 0,
                                            'SLIDER_DEPTH': 0,
                                            'PROGRESS_DEPTH': 0}
    
    sg.LOOK_AND_FEEL_TABLE['tema_escuro'] = {'BACKGROUND': '#000000',
                                             'TEXT': '#F8F8F8',
                                             'INPUT': '#000000',
                                             'TEXT_INPUT': '#F8F8F8',
                                             'SCROLL': '#000000',
                                             'BUTTON': ('#F8F8F8', '#000000'),
                                             'PROGRESS': ('#fca400', '#2A2A2A'),
                                             'BORDER': 0,
                                             'SLIDER_DEPTH': 0,
                                             'PROGRESS_DEPTH': 0}
    
    # define o tema baseado no tema que estava selecionado na última vêz que o programa foi fechado
    f = open(dados_modo, 'r', encoding='utf-8')
    modo = f.read()
    sg.theme(f'tema_{modo}')
    
    def cria_layout():
        coluna_ui = [
            [sg.Button('AJUDA'),
             sg.Button('SOBRE'),
             sg.Button('LOG DO SISTEMA', disabled=True),
             sg.Text('', expand_x=True),
             sg.Button('', key='-tema-', font=("Arial", 11), border_width=0)],
            [sg.Text('')],
            
            [sg.pin(sg.Text(text='Pasta selecionada para capturar os PDFs:', key='-abrir_pdf_texto-', visible=False)),
             sg.pin(RoundedButton(button_text='Selecione a pasta que contenha os arquivos PDF para analisar:', key='-abrir_pdf-', corner_radius=0.8, button_color=('#F8F8F8', '#848484'),
                                  size=(1300, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-input_dir-')),
             sg.pin(sg.InputText(key='-input_dir-', text_color='#fca400', expand_x=True), expand_x=True)],
            
            [sg.pin(sg.Text(text='Pasta selecionada para salvar os e-books:', key='-abrir_pdf_final_texto-', visible=False)),
             sg.pin(RoundedButton(button_text='Selecione a pasta para salvar os e-books:', key='-abrir_pdf_final-', corner_radius=0.8, button_color=('#F8F8F8', '#848484'),
                                  size=(900, 100), button_type=BUTTON_TYPE_BROWSE_FOLDER, target='-output_file-')),
             sg.pin(sg.InputText(key='-output_file-', text_color='#fca400', expand_x=True), expand_x=True)],
            [sg.Text('', expand_y=True)],
            
            [sg.Text('', key='-Mensagens-')],
            [sg.Text(text='', key='-Progresso_texto-'),
             sg.ProgressBar(max_value=0, orientation='h', size=(5, 5), expand_x=True, key='-progressbar-', visible=False)],
            [sg.pin(RoundedButton('INICIAR', key='-iniciar-', corner_radius=0.8, button_color=('#F8F8F8', '#fca400'))),
             sg.pin(RoundedButton('ENCERRAR', key='-encerrar-', corner_radius=0.8, button_color=('#F8F8F8', '#fca400'), visible=False)),
             sg.pin(RoundedButton('ABRIR RESULTADOS', key='-abrir_resultados-', corner_radius=0.8, button_color=('#F8F8F8', '#fca400'), visible=False))]
            ]
        
        coluna_terminal = [[RoundedButton(button_text='TERMINAL ATIVADO', corner_radius=0.8, key='-dev_mode-', button_color=('black', 'red')),
                            sg.Button('🗑', key='-limpa_console-', font=("Arial", 11))],
                           [sg.Output(expand_x=True, expand_y=True, key='-console-')]]
        
        layout = [[sg.Column(coluna_ui, expand_y=True, expand_x=True), sg.Column(coluna_terminal, expand_y=True, expand_x=True, key='-terminal-', visible=False)]]
        
        return sg.Window('Cria E-book DIRPF', layout, finalize=True, resizable=True, return_keyboard_events=True, use_default_focus=False, margins=(30, 30))
        
    
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
        
        # apaga qualquer mensagem na interface
        window['-Mensagens-'].update('')
        
        window['-tema-'].update(disabled=True)
        # habilita e desabilita os botões conforme necessário
        for key in [('-iniciar-', False), ('-encerrar-', True), ('-abrir_resultados-', True), ('-abrir_pdf-', False), ('-abrir_pdf_final-', False),
                    ('-abrir_pdf_texto-', True), ('-abrir_pdf_final_texto-', True), ('-progressbar-', True)]:
            window[key[0]].update(visible=key[1])
            
        # controle para saber se os botões estavam visíveis ao trocar o tema da janela
        with open(controle_botoes, 'w', encoding='utf-8') as f:
            f.write('visible')
        
        try:
            # Chama a função que executa o script
            run(window, pasta_inicial, pasta_final)
            # Qualquer erro o script exibe um alerta e salva gera o arquivo log de erro
        except Exception as erro:
            # Obtém a pilha de chamadas de volta como uma string
            traceback_str = traceback.format_exc()
            escreve_doc(f'Traceback: {traceback_str}\n\n'
                        f'Erro: {erro}')
            window['Log do sistema'].update(disabled=False)
            alert(text='Erro detectado, clique no botão "Log do sistema" para acessar o arquivo de erros e contate o desenvolvedor')
            
        # habilita e desabilita os botões conforme necessário
        for key in [('-encerrar-', False), ('-iniciar-', True), ('-abrir_pdf-', True), ('-abrir_pdf_final-', True),
                    ('-abrir_pdf_texto-', False), ('-abrir_pdf_final_texto-', False), ('-progressbar-', False)]:
            window[key[0]].update(visible=key[1])
        
        window['-tema-'].update(disabled=False)
        window['-Progresso_texto-'].update('')
        window['-Mensagens-'].update('')
    
    
    window = cria_layout()
    window.set_min_size((700, 300))
    
    code = ''
    while True:
        # Obtenha o widget Tkinter subjacente para o elemento Output
        output_widget = window['-console-'].Widget
        
        # configura o tema conforme o tema que estava quando o programa foi fechado da última ves
        f = open(dados_modo, 'r', encoding='utf-8')
        modo = f.read()
        if modo == 'claro':
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-abrir_pdf-', '-abrir_pdf_final-']:
                window[key].update(button_color=('#F8F8F8', None))
            window['-tema-'].update(text='☀', button_color=('#FFC100', '#F8F8F8'))
            # Ajuste a borda diretamente no widget Tkinter
            output_widget.configure(highlightthickness=1, highlightbackground="black")
        else:
            for key in ['-abrir_resultados-', '-encerrar-', '-iniciar-', '-abrir_pdf-', '-abrir_pdf_final-']:
                window[key].update(button_color=('#000000', None))
            window['-tema-'].update(text='🌙', button_color=('#00C9FF', '#000000'))
            # Ajuste a borda diretamente no widget Tkinter
            output_widget.configure(highlightthickness=1, highlightbackground="white")
            
        # captura o evento e os valores armazenados na interface
        event, values = window.read()
        # salva o estado da interface
        save_values(values)
        
        # troca o tema
        if event == '-tema-':
            f = open(dados_modo, 'r', encoding='utf-8')
            modo = f.read()
            
            if modo == 'claro':
                sg.theme('tema_escuro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('escuro')
            else:
                sg.theme('tema_claro')  # Define o tema claro
                with open(dados_modo, 'w', encoding='utf-8') as f:
                    f.write('claro')
            
            window.close()  # Fecha a janela atual
            # inicia as variáveis das janelas
            window = cria_layout()
            
            # retorna os elementos preenchidos
            values = load_values()
            for key, value in values.items():
                if key in window.AllKeysDict:
                    try:
                        window[key].update(value)
                    except Exception as e:
                        print(f"Erro ao atualizar o elemento {key}: {e}")
            
            window['-abrir_pdf-'].update('Selecione a pasta que contenha os arquivos PDF para analisar:')
            window['-abrir_pdf_final-'].update('Selecione a pasta para salvar os e-books:')
                
            # recupera o estado do botão de abrir resultados
            cb = open(controle_botoes, 'r', encoding='utf-8').read()
            if cb == 'visible':
                window['-abrir_resultados-'].update(visible=True)
        
        try:
            pasta_inicial = values['-input_dir-']
            pasta_final = values['-output_file-']
        except:
            pasta_inicial = None
            pasta_final = None
        
        if event == sg.WIN_CLOSED:
            break
        
        elif event == 'Log do sistema':
            os.startfile('Log')
        
        elif event == 'AJUDA':
            os.startfile('Docs\Manual do usuário - Cria E-book DIRPF.pdf')
        
        elif event == 'SOBRE':
            os.startfile('Docs\Sobre.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
    
        elif event == '-abrir_resultados-':
            os.startfile(pasta_final)
        
        elif event == '-limpa_console-':
            window['-console-'].update('')
          
        elif event == '-dev_mode-':
            window['-terminal-'].update(visible=False)
            code = ''
        
        elif event == 'Control_L:17':
            code = ''

        if event == 'Up:38' or event == 'Down:40' or event == 'Left:37' or event == 'Right:39' or event == 'a' or event == 'b':
            code += event
            
        else:
            code = ''
        
        if code == 'Up:38Up:38Down:40Down:40Left:37Right:39Left:37Right:39ab':
            window['-terminal-'].update(visible=True)
            code = ''
            
    window.close()
    