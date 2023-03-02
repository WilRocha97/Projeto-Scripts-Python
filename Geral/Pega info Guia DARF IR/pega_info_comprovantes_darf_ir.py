# -*- coding: utf-8 -*-
import time
import fitz, re, os
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

# Definir os padrões de regex

def analiza():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arquivo in os.listdir(documentos):
        # print(f'\nArquivo: {arquivo}')
        # Abrir o pdf
        arq = os.path.join(documentos, arquivo)
        
        with fitz.open(arq) as pdf:

            # Para cada página do pdf
            for count, page in enumerate(pdf):
                try:
                    # Pega o texto da pagina
                    textinho = page.get_text('text', flags=1 + 2 + 8)
                    
                    #if arquivo == 'Comprovante_Pagamento_02_2022_DARF_007162228638483950 - 21593211000155.pdf':
                        #print(textinho)
                        #return False
                    # Procura codigo 0561
                    # valor_receita = re.compile(r'0561\n.+\n(.+,\d\d)').search(textinho)
                    # if valor_receita:
                    valor = re.compile(r'Totais\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)').search(textinho)
                    if valor:
                        data = re.compile(r'Referência\n(\d\d/\d\d/\d\d\d\d)').search(textinho)
                        empresa = re.compile(r'Data de Vencimento\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)(.+)\n(\d\d/\d\d/\d\d\d\d)').search(textinho)
                        if not empresa:
                            empresa = re.compile(r'Data de Vencimento\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)(.+)\n(\d\d/\d\d\d\d)').search(textinho)
                        
                        valor_principal = valor.group(1)
                        valor_multa = valor.group(2)
                        valor_juros = valor.group(3)
                        valor_total = valor.group(4)
                        
                        data_hora = data.group(1)
    
                        cnpj = empresa.group(1)
                        nome = empresa.group(2)
                        apuracao = empresa.group(3)
    
                        _escreve_relatorio_csv(f"{cnpj};{nome};{apuracao};{valor_principal};{valor_multa};{valor_juros};{valor_total};{data_hora}", nome='Info Comprovantes DARF IR')

                except():
                    print(textinho)
                    

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ;NOME;APURAÇÃO;VALOR PRINCIPAL;VALOR MULTA;VALOR JUROS;VALOR TOTAL;DATA DE ARRECADAÇÃO', nome='Info Comprovantes DARF IR')
    print(datetime.now() - inicio)
