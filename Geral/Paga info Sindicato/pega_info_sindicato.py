# -*- coding: utf-8 -*-
import time
import fitz, re, os
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime
from pathlib import Path


def escreve_relatorio_csv(texto, local, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, "Relatório de Sindicatos.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, "Relatório de Sindicatos - erro.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()


def escreve_header_csv(texto, local, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    with open(os.path.join(local, 'Relatório de Sindicatos.csv'), 'r', encoding=encode) as f:
        conteudo = f.read()
    
    with open(os.path.join(local, 'Relatório de Sindicatos.csv'), 'w', encoding=encode) as f:
        f.write(texto + '\n' + conteudo)


def ask_for_dir():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title='Selecione onde salvar a planilha',
    )
    
    return folder if folder else False


def ask_for_file():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title='Selecione o Relatório de Sindicatos',
        filetypes=[('PDF files', '*.pdf *')],
        initialdir=os.getcwd()
    )
    
    return file if file else False


def guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas, valor_calculado_s, valor_calculado_e):
    codigo_anterior = codigo
    cnpj_anterior = cnpj
    nome_empresa_anterior = nome_empresa
    competencia_anterior = competencia
    sindicato_anterior = sindicato
    nome_anterior = nome
    totais_rubricas_anterior = totais_rubricas
    valor_calculado_s_anterior = valor_calculado_s
    valor_calculado_e_anterior = valor_calculado_e
    
    return codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, nome_anterior, totais_rubricas_anterior, valor_calculado_s_anterior, valor_calculado_e_anterior
    

def analiza():
    # pergunta qual PDF analizar
    sindicato = ask_for_file()
    # pergunta onde salvar a planilha com as informações
    final = ask_for_dir()
    
    # Abrir o pdf
    with fitz.open(sindicato) as pdf:
        
        #inicia as variáveis de armazenamento
        codigo_anterior = ''
        cnpj_anterior = ''
        nome_empresa_anterior = ''
        competencia_anterior = ''
        sindicato_anterior = ''
        nome_anterior = ''
        totais_rubricas_anterior = ''
        valor_calculado_s_anterior = ''
        valor_calculado_e_anterior = ''
        
        # Para cada página do pdf
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                # procura quantas rubricas existem na página
                totais_rubrica = re.compile(r'(Total da Rubrica):\n(.+)').findall(textinho)
                
                # para cada rubrica executa
                for total in totais_rubrica:
                    totais_rubricas = total[1]
                    
                    # procura o nome referênte ao valor da rubrica
                    totais = None
                    indice = 5
                    while not totais:
                        indice = str(indice)
                        totais = re.compile(r'(\d.+ - .+)\nEmpregados\n(.+\n){' + indice + '}(.+)\nTotal da Rubrica:\n(' + totais_rubricas + ')').search(textinho)
                        indice = int(indice)
                        indice += 1
                        
                    nome = totais.group(1)
                    '''print(textinho)
                    time.sleep(555)'''
                    
                    # pega CNPJ, CPF ou CEI da empresa
                    cnpj = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)\nRubrica').search(textinho)
                    if not cnpj:
                        cnpj = re.compile(r'Local de trabalho\n.+\n(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(textinho)
                        if not cnpj:
                            cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d\d)\nRubrica').search(textinho)
                    
                    cnpj = cnpj.group(1)
                    
                    # pega infos da empresa e do sindicato, com algumas variações
                    infos = re.compile(r'Local de trabalho\n(\d) - (.+)\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                    if not infos:
                        infos = re.compile(r'Local de trabalho\n(\d\d) - (.+)\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                        if not infos:
                            infos = re.compile(r'Local de trabalho\n(\d\d\d) - (.+)\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                            if not infos:
                                infos = re.compile(r'Local de trabalho\n(\d\d\d\d) - (.+)\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                                if not infos:
                                    infos = re.compile(r'Local de trabalho\n(\d) - (.+)\n.+\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                                    if not infos:
                                        infos = re.compile(r'Local de trabalho\n(\d\d) - (.+)\n.+\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                                        if not infos:
                                            infos = re.compile(r'Local de trabalho\n(\d\d\d) - (.+)\n.+\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)
                                            if not infos:
                                                infos = re.compile(r'Local de trabalho\n(\d\d\d\d) - (.+)\n.+\n(\d\d/\d\d\d\d)\n.+\n.+\n(.+)').search(textinho)

                    codigo = infos.group(1)
                    nome_empresa = infos.group(2)
                    competencia = infos.group(3)
                    sindicato = infos.group(4)
                    
                    '''if sindicato == 'Sindicato:':
                        print(textinho)
                        time.sleep(333)'''
                    
                    # pega o valor total do sindicato
                    totais_do_sindicato = re.compile(r'(.+)\nTotal do Sindicato:\n(.+)').search(textinho)
                    valor_calculado_s = totais_do_sindicato.group(2)
                    # valor_informado_s = totais_do_sindicato.group(1)
                    
                    # verifica se existe o valor final da empresa na página
                    try:
                        totais_da_empresa = re.compile(r'Total da empresa:\n(.+)\n(.+)').search(textinho)
                        valor_calculado_e = totais_da_empresa.group(1)
                        # valor_informado = totais_da_empresa.group(2)
                    except:
                        valor_calculado_e = ''
                        # valor_informado = 0
                    
                    # se for a primeira página, armazena as infos da empresa
                    if nome_empresa_anterior == '':
                        codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, \
                            nome_anterior, totais_rubricas_anterior, valor_calculado_s_anterior, valor_calculado_e_anterior\
                            = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas, valor_calculado_s, valor_calculado_e)
                        continue

                    # se a emrpresa atual difere da anterior, escreve na planilha as infos coletas da anterior junto dos totais
                    if nome_empresa != nome_empresa_anterior:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};"
                                              f"{competencia_anterior};{valor_calculado_e_anterior};{sindicato_anterior};{valor_calculado_s_anterior};{nome_anterior};{totais_rubricas_anterior}", local=final)

                    # se o sindicato atual difere da anterior, escreve na planilha as infos coletas da anterior junto dos totais do sindicato
                    elif sindicato != sindicato_anterior:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};"
                                              f"{competencia_anterior};;{sindicato_anterior};{valor_calculado_s_anterior};{nome_anterior};{totais_rubricas_anterior}", local=final)

                    # se for a mesma empresa e o mesmo sindicato apenas escreve as infos apenas com o total da rubrica
                    else:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};"
                                              f"{competencia_anterior};;{sindicato_anterior};;{nome_anterior};{totais_rubricas_anterior}", local=final)
                    
                    # armazena as infos da empresa atual
                    codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, \
                        nome_anterior, totais_rubricas_anterior, valor_calculado_s_anterior, valor_calculado_e_anterior \
                        = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas, valor_calculado_s, valor_calculado_e)
                    
            except():
                print(f'\nArquivo: {arq} - ERRO')
        
        # escreve as infos da última página
        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};{competencia_anterior};{sindicato_anterior};"
                              f"{nome_anterior};{totais_rubricas_anterior};{valor_calculado_s_anterior};{valor_calculado_e_anterior}", local=final)
        
        return final


if __name__ == '__main__':
    inicio = datetime.now()
    final = analiza()
    # escreve o cabeçalho da planilha
    escreve_header_csv(';'.join(['CÓDIGO', 'CNPJ', 'NOME', 'COMPETÊNCIA', 'TOTAL EMPRESA', 'SINDICATO', 'TOTAL SINDICATO',
                                 'RUBRICA', 'TOTAL RUBRICA']), local=final)
