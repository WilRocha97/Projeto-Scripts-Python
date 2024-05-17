# -*- coding: utf-8 -*-
import time, shutil, io, fitz, PyPDF2, re, os, sys, traceback, PySimpleGUI as sg
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

# Carregar a fonte TrueType (substitua 'sua_fonte.ttf' pelo caminho da sua fonte)
pdfmetrics.registerFont(TTFont('Fonte', 'Assets\HankenGrotesk-SemiBold.ttf'))
tamanho_da_pagina = (672,950)


def data_modificacao_arquivo(nome_arquivo):
    # Obter o tempo de modificação do arquivo
    tempo_modificacao = os.path.getmtime(nome_arquivo)
    return tempo_modificacao


# Caminho dos dois arquivos que você deseja comparar
caminho_arquivo1 = 'C:\Program Files (x86)\Automações\Cria E-book DIRPF\cria_e-book.exe'
caminho_arquivo2 = 'T:\ROBÔ\_Executáveis\Cria E-book\Cria E-book DIRPF.exe'
data_modificacao1 = 0
data_modificacao2 = 0

try:
    # Obter a data de modificação de cada arquivo
    data_modificacao1 = data_modificacao_arquivo(caminho_arquivo1)
    data_modificacao2 = data_modificacao_arquivo(caminho_arquivo2)
    
    data_modificacao1 = int(str(data_modificacao1)[:7])
    data_modificacao2 = int(str(data_modificacao2)[:7])
    
    print(data_modificacao1)
    print(data_modificacao2)
except:
    pass
    
# Comparar as datas de modificação
if data_modificacao1 < data_modificacao2:
    while True:
        atualizar = confirm(text='Existe uma nova versão do programa, deseja atualizar agora?', buttons=('Novidades da atualização', 'Sim', 'Não'))
        if atualizar == 'Novidades da atualização':
            os.startfile('Sobre.pdf')
        if atualizar == 'Sim':
            break
        elif atualizar == 'Não':
            break
    
    if atualizar == 'Sim':
        os.startfile(caminho_arquivo2)
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
    
    pdf_merger.append(os.path.join('Assets', 'AGRADECIMENTO POR ESCOLHER A VEIGA & POSTAL - IPRF TIMBRADO.pdf'))

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
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    # Carregar a fonte personalizada
    sg.theme()
    
    sg.LOOK_AND_FEEL_TABLE['tema'] = {'BACKGROUND': '#ffffff',
                                      'TEXT': '#0e0e0e',
                                      'INPUT': '#ffffff',
                                      'TEXT_INPUT': '#0e0e0e',
                                      'SCROLL': '#ffffff',
                                      'BUTTON': ('#0e0e0e', '#ffffff'),
                                      'PROGRESS': ('#ffffff', '#ffffff'),
                                      'BORDER': 0,
                                      'SLIDER_DEPTH': 0,
                                      'PROGRESS_DEPTH': 0}
    
    sg.theme('tema')  # Define o tema do PySimpleGUI
    layout = [
        [sg.Button('AJUDA', font=("Helvetica", 10, "underline"), border_width=0),
         sg.Button('SOBRE', font=("Helvetica", 10, "underline"), border_width=0),
         sg.Button('LOG DO SISTEMA', font=("Helvetica", 10, "underline"), border_width=0, disabled=True)],
        [sg.Text('')],
        
        [sg.Text('Selecione a pasta que contenha os arquivos PDF para analisar:')],
        [sg.FolderBrowse('SELECIONAR', font=("Helvetica", 10, "underline"), key='-abrir_pdf-'),
         sg.InputText(key='-output_dir-', size=80, disabled=True)],
        [sg.Text('')],
        
        [sg.Text('Selecione a pasta para salvar os e-books:')],
        [sg.FolderBrowse('SELECIONAR', font=("Helvetica", 10, "underline"), key='-abrir_pdf_final-'),
         sg.InputText(key='-output_dir-', size=80, disabled=True)],
        [sg.Text('', expand_y=True)],
        
        [sg.Text('', key='-Mensagens-')],
        [sg.Text(size=6, text='', key='-Progresso_texto-'),
         sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color=('#fca400', '#ffe0a6'), visible=False)],
        [sg.Button('INICIAR', font=("Helvetica", 10, "underline"), button_color=('white', 'white'), image_filename='Assets\\fundo_botao.png', key='-iniciar-', border_width=0),
         sg.Button('ENCERRAR', font=("Helvetica", 10, "underline"), key='-encerrar-', disabled=True, border_width=0),
         sg.Button('ABRIR RESULTADOS', font=("Helvetica", 10, "underline"), key='-abrir_resultados-', disabled=True, border_width=0)],
    ]

    # guarda a janela na variável para manipula-la
    window = sg.Window('Cria E-book DIRPF', layout, finalize=True, resizable=True, margins=(30, 30))
    window.set_min_size((500, 300))
    
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
        window['-progressbar-'].update(visible=True)
        
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
            
        window['-progressbar-'].update_bar(0)
        window['-progressbar-'].update(visible=False)
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
        
        elif event == 'Sobre':
            os.startfile('Sobre.pdf')
        
        elif event == '-iniciar-':
            # Cria uma nova thread para executar o script
            script_thread = Thread(target=run_script_thread)
            script_thread.start()
    
        elif event == '-abrir_resultados-':
            os.startfile(pasta_final)
    
    window.close()
    