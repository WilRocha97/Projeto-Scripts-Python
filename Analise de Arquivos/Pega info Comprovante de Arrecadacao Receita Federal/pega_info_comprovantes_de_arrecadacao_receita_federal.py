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
        
        try:
            with fitz.open(arq) as pdf:
    
                # Para cada página do pdf
                for count, page in enumerate(pdf):
                    try:
                        # Pega o texto da pagina
                        textinho = page.get_text('text', flags=1 + 2 + 8)
                        
                        print(textinho)
                        
                        #if arquivo == 'Comprovante_Pagamento_02_2022_DARF_007162228638483950 - 21593211000155.pdf':
                            #print(textinho)
                            #return False
                        # Procura codigo 0561
                        # valor_receita = re.compile(r'0561\n.+\n(.+,\d\d)').search(textinho)
                        # if valor_receita:
                        valores = re.compile(r'\n(\d\d\d\d)\n(\w.+)\n(.+)\n(.+)\n(.+)\n(.+)').findall(textinho)
                        if valores:
                            for valor in valores:
                                empresa = re.compile(r'Data de Vencimento\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)(.+)\n(\d\d/\d\d/\d\d\d\d)\n.+\n(.+)').search(textinho)
                                if not empresa:
                                    empresa = re.compile(r'Data de Vencimento\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)(.+)\n(\d\d/\d\d\d\d)\n.+\n(.+)').search(textinho)
                                
                                codigo = valor[0]
                                descricao = valor[1]
                                if descricao == 'Totais':
                                    continue
                                valor_principal = valor[2]
                                valor_multa = valor[3]
                                valor_juros = valor[4]
                                valor_total = valor[5]
            
                                cnpj = empresa.group(1)
                                nome = empresa.group(2)
                                apuracao = empresa.group(3)
                                num_comprovante = empresa.group(4)
                                
                                banco = re.compile('(\d\d/\d\d/\d\d\d\d)\n(Documento pago via PIX)').search(textinho)
                                if not banco:
                                    banco = re.compile('(\d\d/\d\d/\d\d\d\d)\n(.+)\n(.+)\n(.+)\n(\d+,\d+)').search(textinho)
                                    if not banco:
                                        banco = re.compile('(\d\d/\d\d/\d\d\d\d)\n(.+)\n(.+)\n(\d+,\d+)').search(textinho)
                                        banco_nome = banco.group(2)
                                        arrecadacao = banco.group(1)
                                        agencia = 'x'
                                        estabelecimento = banco.group(3)
                                        val_restituido = banco.group(4)
                                    else:
                                        banco_nome = banco.group(2)
                                        arrecadacao = banco.group(1)
                                        agencia = banco.group(3)
                                        estabelecimento = banco.group(4)
                                        val_restituido = banco.group(5)
                                    
                                else:
                                    banco_nome = banco.group(2)
                                    arrecadacao = banco.group(1)
                                    agencia = ''
                                    estabelecimento = ''
                                    val_restituido = ''
                                
                                _escreve_relatorio_csv(f"{cnpj};{nome};{apuracao};'{num_comprovante};"
                                                       f"{codigo};{descricao};{valor_principal};{valor_multa};"
                                                       f"{valor_juros};{valor_total};{banco_nome};{arrecadacao};"
                                                       f"{agencia};{estabelecimento};{val_restituido}", nome='Info Comprovantes')
    
                    except():
                        print(textinho)
        except:
            print(f'{arq} Não é um arquivo válido')
            _escreve_relatorio_csv(f'{arq};Arquivo corrompido', nome='Info Comprovantes')

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('CNPJ;NOME;APURAÇÃO;NÚMERO DO DOCUMENTO;CÓDIGO;DESCRIÇÃO;VALOR PRINCIPAL;VALOR MULTA;VALOR JUROS;VALOR TOTAL;BANCO;DATA DE ARRECADAÇÃO;AGÊNCIA;ESTABELECIMENTO;VALOR RESTITUÍDO', nome='Info Comprovantes')
    print(datetime.now() - inicio)
