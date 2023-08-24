# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_file
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv


def analiza():
    arq = ask_for_file(filetypes=[('PDF files', '*.pdf *')])
    # Abrir o pdf
    with fitz.open(arq) as pdf:
        # Para cada página do pdf, se for a segunda página o script ignora
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                empresa = re.compile(r'(\d+) - (\w.+)\n(.+)').search(textinho)
                
                try:
                    cod_empresa = empresa.group(1)
                    nome = empresa.group(2)
                    cnpj = empresa.group(3)
                
                except:
                    _escreve_relatorio_csv(f'{page.number};erro empresa', nome='erros')
                
                codigos = re.compile(r'(\d\d\d\d)\n(\w+)\n(.+,.+)\n(\d\d/\d\d\d\d)\n(.+,.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)').findall(textinho)
                if not codigos:
                    pass
                
                else:
                    for codigo in codigos:
                        cod = codigo[0]
                        periodicidade = codigo[1]
                        periodo = codigo[3]
                        val_acumulado = codigo[4]
                        val_recolher = codigo[5]
                        val_compensar = codigo[6]
                        val_pagar = codigo[7]
                        val_acumular = codigo[2]
                        
                        _escreve_relatorio_csv(f'{cod_empresa};{nome};{cnpj};{cod};{periodicidade};{periodo};{val_acumulado};{val_recolher};{val_compensar};{val_pagar};{val_acumular}')
            
            except():
                print(textinho)
                print('ERRO')
            
            return periodo
            

if __name__ == '__main__':
    inicio = datetime.now()
    periodo = analiza()
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;PERIODICIDADE;PERÍODO;CÓDIGO RECOLHIMENTO;VALOR A RECOLHER;VALOR A COMPENSAR;VALOR A PAGAR;VALOR A ACUMULAR', nome=f'Encargos de IRRF - {periodo.replace("/", "-")}')
    print(datetime.now() - inicio)
