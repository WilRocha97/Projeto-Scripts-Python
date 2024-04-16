# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime
from pyautogui import confirm
from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

def analiza():
    documentos = ask_for_dir()
    analise_completa = confirm(buttons=('Analise completa', 'Apenas CNPJ e valor total da guia'))
    # Analiza cada pdf que estiver na pasta
    for arq_nome in os.listdir(documentos):
        if not arq_nome.endswith('.pdf'):
            continue
        print(f'\nArquivo: {arq_nome}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq_nome)
        with fitz.open(arq) as pdf:
            tem_imposto = False
            textinho = ''
            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                texto_pagina = page.get_text('text', flags=1 + 2 + 8)
                textinho += texto_pagina
            
            try:
                textinho_parte_1 = textinho.split('SENDA')[0]
                textinho_parte_2 = textinho.split('SENDA')[1].split('Composição do Documento de Arrecadação')[1]
                textinho = textinho_parte_1 + textinho_parte_2
            except:
                pass
            
            # Procura o valor a recolher da empresa
            try:
                cnpj = re.compile(r'Documento de Arrecadação\nde Receitas Federais\n(\d.+)\n').search(textinho).group(1)
            except:
                try:
                    cnpj = re.compile(r'Documento de Arrecadação\ndo eSocial\n(\d.+)\n').search(textinho).group(1)
                except:
                    print('Erro')
                    _escreve_relatorio_csv(f'{arq_nome}', nome='Arquivos com erro')
                    continue
            
            # Procura a descrição do valor a recolher 1, tem algumas variações do que aparece junto a essa info
            try:
                valor = re.compile(r'Valor Total do Documento\n(.+)\nCNPJ').search(textinho).group(1)
            except:
                valor = re.compile(r'Valor Total do Documento\n(.+)\nCPF').search(textinho).group(1)
            
            if analise_completa == 'Analise completa':
                regexes = [r'(1082)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1082)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1082)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1099)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1099)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1099)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1138)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1138)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1138)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1170)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1170)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1170)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1176)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1176)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1176)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1181)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1181)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1181)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1184)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1184)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1184)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1200)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1200)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1200)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1646)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1646)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1646)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(1708)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1708)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(1708)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)',
                           r'(5952)\n(.+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(5952)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)', r'(5952)\n(.+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n.+\n(.+)']
                for regex in regexes:
                    imposto = re.compile(regex).search(textinho)
                    if imposto:
                        try:
                            imposto_capturado = (f'{imposto.group(1)};{imposto.group(2)};{imposto.group(3)};{imposto.group(4)};'
                                                 f'{imposto.group(5)};{imposto.group(6)};{imposto.group(7)}')
                        except:
                            try:
                                imposto_capturado = (f'{imposto.group(1)};{imposto.group(2)};{imposto.group(3)};0;'
                                                    f'{imposto.group(4)};{imposto.group(5)};{imposto.group(6)}')
                            except:
                                imposto_capturado = (f'{imposto.group(1)};{imposto.group(2)};{imposto.group(3)};0;0;{imposto.group(4)};'
                                                     f'{imposto.group(5)}')
                        
                        print(f'{cnpj} - {valor}')
                        _escreve_relatorio_csv(f"{cnpj};{imposto_capturado};{valor}")
                        tem_imposto = True
            
            if not tem_imposto:
                _escreve_relatorio_csv(f"{cnpj};{valor}")
                # print(textinho)
    
    if analise_completa == 'Analise completa':
        _escreve_header_csv(';'.join(['CNPJ', 'CÓDIGO', 'IMPOSTO', 'PRINCIPAL', 'MULTA', 'JUROS', 'TOTAL IMPOSTO', 'COMPETÊNCIA / VENCIMENTO', 'VALOR TOTAL GUIA']))
    if analise_completa == 'Apenas CNPJ e valor total da guia':
        _escreve_header_csv('CNPJ;VALOR TOTAL GUIA')

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    
    print(datetime.now() - inicio)