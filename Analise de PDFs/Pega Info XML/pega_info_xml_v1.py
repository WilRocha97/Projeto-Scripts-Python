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

            arq = BeautifulSoup(data, "xml")
            arq = arq.prettify()
            # print(arq)
            # time.sleep(55)

            pattern = re.compile(r'\n\n+')
            arq = re.sub(pattern, '', arq)
            
            notas = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ (.+)\n.+ </ns3:Numero>').findall(arq)
            if not notas:
                print(f'Erro ao ler documento: {arq_name}, texto do arquivo não confere com o layout esperado.')
                continue
            
            for nota in notas:
                _escreve_relatorio_csv(f'{nota}', nome='Notas')
                
                data_emissao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){4}\n.+<ns3:DataEmissao>\n.+ (.+)').search(arq).group(2)
                competencia = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){19}\n.+<ns3:Competencia>\n.+ (.+)').search(arq).group(2)
                
                try:
                    nota_substituida = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){23}\n.+<ns3:Numero>\n.+ (.+)').search(arq).group(2)
                except:
                    nota_substituida = 'Original'
                
                # Orden das tags: com nota subtituída, com pis confins, com nota substituída e pis cofins
                tags = [24, 32]
                for tag in range(50):
                    try:
                        valor_servico = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:ValorServicos>\n.+ (.+)').search(arq).group(2)
                    except:
                        pass

                tags = [30, 38, 42, 50]
                for tag in range(60):
                    try:
                        valor_iss = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:ValorIss>\n.+ (.+)').search(arq)
                        if valor_iss:
                            break
                    except:
                        pass
                valor_iss = valor_iss.group(2)

                tags = [33, 41, 45, 53]
                for tag in range(60):
                    try:
                        base_de_calculo = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:BaseCalculo>\n.+ (.+)').search(arq)
                        if base_de_calculo:
                            break
                    except:
                        pass
                base_de_calculo = base_de_calculo.group(2)

                tags = [36, 44, 48, 56]
                for tag in range(60):
                    try:
                        aliquota = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:Aliquota>\n.+ (.+)').search(arq)
                        if aliquota:
                            break
                    except:
                        pass
                aliquota = aliquota.group(2)

                tags = [39, 47, 51, 59]
                for tag in range(70):
                    try:
                        valor_liquido = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:ValorLiquidoNfse>\n.+ (.+)').search(arq)
                        if valor_liquido:
                            break
                    except:
                        pass
                valor_liquido = valor_liquido.group(2)

                tags = [52, 60, 64, 72]
                for tag in range(80):
                    try:
                        discriminacao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:Discriminacao>\n.+ {2}(.+)').search(arq)
                        if discriminacao:
                            break
                    except:
                        pass
                discriminacao = discriminacao.group(2)

                tags = [61, 69, 73, 74, 81]
                for tag in range(100):
                    try:
                        cnpj_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:Cnpj>\n.+ (.+)').search(arq)
                        if cnpj_prestador:
                            break
                    except:
                        pass
                try:
                    cnpj_prestador = cnpj_prestador.group(2)
                except:
                    print(tag)
                    print(nota)
                    print(arq)
                    
                tags = [68, 76, 80, 81, 88]
                for tag in range(100):
                    try:
                        nome_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:RazaoSocial>\n.+ {2}(.+)').search(arq)
                        if nome_prestador:
                            break
                    except:
                        pass
                nome_prestador = nome_prestador.group(2)

                tags = [105, 113, 114, 117, 118, 125]
                for tag in range(130):
                    try:
                        cnpj_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:CpfCnpj>\n.+\n.+ (.+)').search(arq)
                        if cnpj_tomador:
                            break
                    except:
                        pass
                cnpj_tomador = cnpj_tomador.group(2)
                
                # Orden das tags: com nota subtituída, com pis confins, com pis confins e tomador cnpj, com nota substituída e pis cofins
                tags = [111, 119, 120, 123, 124, 126, 127, 134]
                for tag in range(140):
                    try:
                        nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){' + str(tag) + '}\n.+<ns3:RazaoSocial>\n.+ {2}(.+)').search(arq)
                        if nome_tomador:
                            break
                    except:
                        pass
                nome_tomador = nome_tomador.group(2)
                
                _escreve_relatorio_csv(f'{nota};{nota_substituida};{data_emissao};{competencia};{cnpj_prestador};{nome_prestador.replace("&amp;", "&")};{cnpj_tomador};{nome_tomador};{discriminacao};'
                                       f'{valor_servico};{valor_liquido};{valor_iss};{base_de_calculo};{aliquota}', nome='Info XML')
            
        
if __name__ == '__main__':
    inicio = datetime.now()
    if not analiza():
        print('')
    else:
        _escreve_header_csv(';'.join(['NOTA', 'NOTA SUBSTITUÍDA', 'EMISSÃO', 'COMPETÊNCIA', 'CNPJ PRESTADOR', 'RAZÃO SOCIAL PRESTADOR', 'CPF / CNPJ TOMADOR',
                                  'RAZÃO SOCIAL TOMADOR', 'DISCRIMINAÇÃO', 'VALOR', 'VALOR LIQUIDO', 'VALOR ISS', 'BASE DE CÁLCULO', 'ALIQUOTA']), nome='Info XML')
    print(datetime.now() - inicio)
