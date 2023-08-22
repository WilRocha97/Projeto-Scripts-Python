# -*- coding: utf-8 -*-
import time
import re, os
import lxml
from bs4 import BeautifulSoup
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv

# Definir os padrões de regex


def analiza():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arq_name in os.listdir(documentos):
        # Abrir o pdf
        arquivo = os.path.join(documentos, arq_name)
        with open(arquivo, 'r', encoding='utf-8') as f:
            data = f.read()

            arq_bs = BeautifulSoup(data, "xml")
            arq = arq_bs.prettify()
            # print(arq)
            # time.sleep(55)
            
            nota = re.compile(r'InfNfse Id=\"(.+)\">\n.+Numero>\n +(\d+)').search(arq)
            id_nota = nota.group(1)
            num_nota = nota.group(2)
            
            data_emissao = re.compile(r'DataEmissao>\n +(.+)\n +</').search(arq).group(1)
            data_emissao = data_emissao.replace('T', '|')
            
            competencia = re.compile(r'Competencia>\n +(.+)\n +</').search(arq).group(1)
            base_calculo = re.compile(r'BaseCalculo>\n +(.+)\n +</').search(arq).group(1)
            aliquota = re.compile(r'BaseCalculo>\n +<.+:Aliquota>\n +(.+)\n +</').search(arq).group(1)
            valor_iss = re.compile(r'Aliquota>\n +<.+:ValorIss>\n +(.+)\n +</').search(arq).group(1)
            valor_liquido = re.compile(r'ValorIss>\n +<.+:ValorLiquidoNfse>\n +(.+)\n +</').search(arq).group(1)
            valor_servico = re.compile(r'ValorServicos>\n +(.+)\n +</').search(arq).group(1)
            
            for i in range(10):
                try:
                    discriminacao = re.compile(r'Discriminacao>\n +((.+\n){' + str(i) + '}) +</ns2:Discriminacao>').search(arq).group(1)
                    discriminacao = discriminacao.replace('\n', '. ').replace('..', '. ').replace(': . ', ': ')
                    discriminacao = re.sub('\s+', ' ', discriminacao)
                    break
                except:
                    discriminacao = ''
                    pass
                    
            nome_tomador = re.compile(r'IdentificacaoTomador>\n +<.+:RazaoSocial>\n +(.+)\n +</').search(arq).group(1)
            try:
                cnpj_tomador = re.compile(r'IdentificacaoTomador>\n +<.+:CpfCnpj>\n +<.+:Cnpj>\n +(.+)\n +</').search(arq).group(1)
            except:
                cnpj_tomador = re.compile(r'IdentificacaoTomador>\n +<.+:CpfCnpj>\n +<.+:Cpf>\n +(.+)\n +</').search(arq).group(1)
            
            nome_prestador = re.compile(r'PrestadorServico>\n +<.+:RazaoSocial>\n +(.+)\n +</').search(arq).group(1)
            try:
                cnpj_prestador = re.compile(r'Prestador>\n +<.+:CpfCnpj>\n +<.+:Cnpj>\n +(.+)\n +</').search(arq).group(1)
            except:
                cnpj_prestador = re.compile(r'Prestador>\n +<.+:CpfCnpj>\n +<.+:Cpf>\n +(.+)\n +</').search(arq).group(1)
                
            _escreve_relatorio_csv(f'{id_nota};{num_nota};{data_emissao};{competencia};{cnpj_prestador};{nome_prestador.replace("&amp;", "&")};{cnpj_tomador};{nome_tomador};{discriminacao};'
                                       f'{valor_servico};{valor_liquido};{valor_iss};{base_calculo};{aliquota}', nome='Info XML')
    
        
if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv('ID DA NOTA;NOTA;EMISSÃO;COMPETÊNCIA;CNPJ PRESTADOR;RAZÃO SOCIAL PRESTADOR;CPF CNPJ TOMADOR;RAZÃO SOCIAL TOMADOR;DISCRIMINAÇÃO;VALOR;VALOR LIQUIDO;VALOR ISS;BASE DE CÁLCULO;ALÍQUOTA', nome='Info XML')

    print(datetime.now() - inicio)
