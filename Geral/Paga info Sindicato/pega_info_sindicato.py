# -*- coding: utf-8 -*-
import time
import fitz, re, os
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime
from pathlib import Path

e_dir = Path('execução')


def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    if local == e_dir:
        local = Path(local)
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(str(local / f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome}-erro.csv"), 'a', encoding=encode)
    
    f.write(texto + end)
    f.close()


def escreve_header_csv(texto, nome='resumo.csv', local=e_dir, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    with open(str(local / nome), 'r', encoding=encode) as f:
        conteudo = f.read()
    
    with open(str(local / nome), 'w', encoding=encode) as f:
        f.write(texto + '\n' + conteudo)


def ask_for_file(title='Abrir arquivo', filetypes='*', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
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
    sindicato = ask_for_file(title='Selecione o Relatório de Sindicatos', filetypes = [('PDF files', '*.pdf *')])
    # Abrir o pdf
    with fitz.open(sindicato) as pdf:
        # Para cada página do pdf
        codigo_anterior = ''
        cnpj_anterior = ''
        nome_empresa_anterior = ''
        competencia_anterior = ''
        sindicato_anterior = ''
        nome_anterior = ''
        totais_rubricas_anterior = ''
        valor_calculado_s_anterior = ''
        valor_calculado_e_anterior = ''
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o valor a recolher da empresa

                totais_rubrica = re.compile(r'(Total da Rubrica):\n(.+)').findall(textinho)
                    
                for total in totais_rubrica:
                    totais_rubricas = total[1]
                    
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
                    
                    cnpj = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)\nRubrica').search(textinho)
                    if not cnpj:
                        cnpj = re.compile(r'Local de trabalho\n.+\n(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(textinho)
                        if not cnpj:
                            cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d\d)\nRubrica').search(textinho)
                    
                    cnpj = cnpj.group(1)
                    
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

                    totais_do_sindicato = re.compile(r'(.+)\nTotal do Sindicato:\n(.+)').search(textinho)
                    valor_calculado_s = totais_do_sindicato.group(2)
                    # valor_informado_s = totais_do_sindicato.group(1)
                    
                    try:
                        totais_da_empresa = re.compile(r'Total da empresa:\n(.+)\n(.+)').search(textinho)
                        valor_calculado_e = totais_da_empresa.group(1)
                        # valor_informado = totais_da_empresa.group(2)
                    except:
                        valor_calculado_e = ''
                        # valor_informado = 0
                    
                    if nome_empresa_anterior == '':
                        codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, \
                            nome_anterior, totais_rubricas_anterior, valor_calculado_s_anterior, valor_calculado_e_anterior\
                            = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas, valor_calculado_s, valor_calculado_e)
                        continue
                        
                    if nome_empresa != nome_empresa_anterior:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};{competencia_anterior};{sindicato_anterior};"
                                              f"{nome_anterior};{totais_rubricas_anterior};{valor_calculado_s_anterior};{valor_calculado_e_anterior}")
                        
                    elif sindicato != sindicato_anterior:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};{competencia_anterior};{sindicato_anterior};"
                                              f"{nome_anterior};{totais_rubricas_anterior};{valor_calculado_s_anterior};{valor_calculado_e_anterior}")
                        
                    else:
                        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};{competencia_anterior};{sindicato_anterior};"
                                              f"{nome_anterior};{totais_rubricas_anterior}")

                    codigo_anterior, cnpj_anterior, nome_empresa_anterior, competencia_anterior, sindicato_anterior, \
                        nome_anterior, totais_rubricas_anterior, valor_calculado_s_anterior, valor_calculado_e_anterior \
                        = guarda_info(codigo, cnpj, nome_empresa, competencia, sindicato, nome, totais_rubricas, valor_calculado_s, valor_calculado_e)
                    
            except():
                print(f'\nArquivo: {arq} - ERRO')
                
        escreve_relatorio_csv(f"{codigo_anterior};{cnpj_anterior};{nome_empresa_anterior};{competencia_anterior};{sindicato_anterior};"
                              f"{nome_anterior};{totais_rubricas_anterior};{valor_calculado_s_anterior};{valor_calculado_e_anterior}")


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    escreve_header_csv(';'.join(['CÓDIGO', 'CNPJ', 'NOME', 'COMPETÊNCIA', 'SINDICATO', 'RUBRICA', 'TOTAL DA RUBRICA', 'TOTAL DO SINDICATO CALCULADO', 'TOTAL DA EMPRESA CALCULADO']))
    print(datetime.now() - inicio)
