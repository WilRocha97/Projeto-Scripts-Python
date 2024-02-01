# -*- coding: utf-8 -*-
import time
import fitz, re, os
import tkinter.filedialog
import pyautogui


def escreve_relatorio_csv(texto, local, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"Relatório de Sindicatos.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, "Relatório de Sindicatos - complementar.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()


def escreve_header_csv(texto, local, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    with open(os.path.join(local, 'Relatório de Sindicatos.csv'), 'r', encoding=encode) as f:
        conteudo = f.read()
    
    with open(os.path.join(local, 'Relatório de Sindicatos.csv'), 'w', encoding=encode) as f:
        f.write(texto + '\n' + conteudo)


def ask_for_dir():
    root = tkinter.filedialog.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = tkinter.filedialog.askdirectory(
        title='Selecione onde salvar a planilha',
    )
    
    return folder if folder else False


def ask_for_file():
    root = tkinter.filedialog.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = tkinter.filedialog.askopenfilename(
        title='Selecione o Relatório de Sindicatos',
        filetypes=[('PDF files', '*.pdf *')],
        initialdir=os.getcwd()
    )
    
    return file if file else False


def guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas):
    codigo_anterior = codigo
    cnpj_anterior = cnpj
    nome_empresa_anterior = nome_empresa
    competencia_anterior = competencia
    sindicato_anterior = sindicato
    nome_anterior = nome
    totais_rubricas_anterior = totais_rubricas
    
    return codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, nome_anterior, totais_rubricas_anterior
    

def guarda_valores_totais(valor_calculado_s, valor_calculado_e):
    valor_calculado_s_anterior = valor_calculado_s
    valor_calculado_e_anterior = valor_calculado_e
    return valor_calculado_s_anterior, valor_calculado_e_anterior


def analiza():
    # pergunta qual PDF analisar
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
                print(page.number - 1)
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                # procura quantas rubricas existem na página
                totais_rubrica = re.compile(r'(Total da Rubrica):\n(.+)').findall(textinho)
                
                if not totais_rubrica:
                    # pega o valor total do sindicato
                    try:
                        totais_do_sindicato = re.compile(r'(.+)\nTotal do Sindicato:\n(.+)').search(textinho)
                        valor_calculado_s = totais_do_sindicato.group(2)
                    except:
                        valor_calculado_s = ''
                    # verifica se existe o valor final da empresa na página
                    try:
                        totais_da_empresa = re.compile(r'Total da empresa:\n(.+)\n(.+)').search(textinho)
                        valor_calculado_e = totais_da_empresa.group(1)
                    except:
                        valor_calculado_e = ''
                    valor_calculado_s_anterior, valor_calculado_e_anterior = guarda_valores_totais(valor_calculado_s, valor_calculado_e)
                
                # para cada rubrica executa
                for total in totais_rubrica:
                    totais_rubricas = total[1]
                    
                    # procura o nome referente ao valor da rubrica
                    totais = None
                    indice = 0
                    
                    while not totais:
                        indice = str(indice)
                        if re.compile(r'(\d.+ - .+)\nEmpregados\n').search(textinho):
                            totais = re.compile(r'(\d.+ - .+)\nEmpregados\n(.+\n){' + indice + '}(.+)\nTotal da Rubrica:\n(' + totais_rubricas + ')').search(textinho)
                        
                        if re.compile(r'(\d.+ - .+)\nContribuintes\n').search(textinho):
                            totais = re.compile(r'(\d.+ - .+)\nContribuintes\n(.+\n){' + indice + '}(.+)\nTotal da Rubrica:\n(' + totais_rubricas + ')').search(textinho)
                        
                        if re.compile(r'(\d.+ - .+)\nEstagiários\n').search(textinho):
                            totais = re.compile(r'(\d.+ - .+)\nEstagiários\n(.+\n){' + indice + '}(.+)\nTotal da Rubrica:\n(' + totais_rubricas + ')').search(textinho)
    
                        indice = int(indice)
                        indice += 1
                        
                    nome = totais.group(1)
                    
                    # pega CNPJ, CPF ou CEI da empresa
                    cnpj = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)\nRubrica').search(textinho)
                    if not cnpj:
                        cnpj = re.compile(r'Local de trabalho\n.+\n(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(textinho)
                        if not cnpj:
                            cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d\d)\nRubrica').search(textinho)
                            if not cnpj:
                                cnpj = re.compile(r'(\d\d\d.\d\d\d.\d\d\d/\d\d\d-\d\d)\n\d+/\d\d/\d\d\d\d').search(textinho)
                    
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
                    try:
                        totais_do_sindicato = re.compile(r'(.+)\nTotal do Sindicato:\n(.+)').search(textinho)
                        valor_calculado_s = totais_do_sindicato.group(2)
                    except:
                        valor_calculado_s = ''
                    # verifica se existe o valor final da empresa na página
                    try:
                        totais_da_empresa = re.compile(r'Total da empresa:\n(.+)\n(.+)').search(textinho)
                        valor_calculado_e = totais_da_empresa.group(1)
                    except:
                        valor_calculado_e = ''
                    
                    # se for a primeira página, armazena as infos da empresa
                    if nome_empresa_anterior == '':
                        codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, \
                            nome_anterior, totais_rubricas_anterior \
                            = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas)
                        valor_calculado_s_anterior, valor_calculado_e_anterior = guarda_valores_totais(valor_calculado_s, valor_calculado_e)
                        continue
                        
                    # se a empresa atual difere da anterior, escreve na planilha as infos coletas da anterior junto dos totais
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
                        nome_anterior, totais_rubricas_anterior \
                        = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas)
                    valor_calculado_s_anterior, valor_calculado_e_anterior = guarda_valores_totais(valor_calculado_s, valor_calculado_e)
                    
                    
            except():
                escreve_relatorio_csv(f"{page.number};Erro", local=final)
        
        # escreve as infos da última página
        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};"
                              f"{competencia_anterior};{valor_calculado_e_anterior};{sindicato_anterior};{valor_calculado_s_anterior};{nome_anterior};{totais_rubricas_anterior}", local=final)
        
        return final


if __name__ == '__main__':
    #try:
    final = analiza()
    # escreve o cabeçalho da planilha
    escreve_header_csv(';'.join(['CÓDIGO', 'CNPJ', 'NOME', 'COMPETÊNCIA', 'TOTAL EMPRESA', 'SINDICATO', 'TOTAL SINDICATO',
                                 'RUBRICA', 'TOTAL RUBRICA']), local=final)
    
    
    pyautogui.alert(title='Gera relatório de sindicato', text='Relatório gerado com sucesso.')
    """except:
        pyautogui.alert(title='Gera relatório de sindicato', text='Cancelado.')"""
