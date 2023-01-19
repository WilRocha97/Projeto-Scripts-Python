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
    for arq in os.listdir(documentos):
        # Abrir o pdf
        arquivo = os.path.join(documentos, arq)
        with open(arquivo, 'r', encoding='utf-8') as f:
            data = f.read()

            arq = BeautifulSoup(data, "xml")
            arq = arq.prettify()
            # print(arq)
            # time.sleep(55)

            try:
                notas = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ (.+)\n.+ </ns3:Numero>').findall(arq)
                
                for nota in notas:
                    data_emissao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){4}\n.+<ns3:DataEmissao>\n.+ (.+)').search(arq).group(2)
                    competencia = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){19}\n.+<ns3:Competencia>\n.+ (.+)').search(arq).group(2)
                    
                    try:
                        nota_substituida = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){23}\n.+<ns3:Numero>\n.+ (.+)').search(arq).group(2)
                    except:
                        nota_substituida = 'Original'
                    
                    try:
                        valor_servico = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){24}\n.+<ns3:ValorServicos>\n.+ (.+)').search(arq).group(2)
                    except:
                        # com nota substituída
                        valor_servico = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){32}\n.+<ns3:ValorServicos>\n.+ (.+)').search(arq).group(2)
                    
                    
                    valor_iss = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){30}\n.+<ns3:ValorIss>\n.+ (.+)').search(arq)
                    if not valor_iss:
                        # com nota substituída
                        valor_iss = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){38}\n.+<ns3:ValorIss>\n.+ (.+)').search(arq)
                        if not valor_iss:
                            # com pis cofins
                            valor_iss = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){42}\n.+<ns3:ValorIss>\n.+ (.+)').search(arq)
                            if not valor_iss:
                                # com nota substituída e pis cofins
                                valor_iss = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){50}\n.+<ns3:ValorIss>\n.+ (.+)').search(arq)
                    
                    valor_iss = valor_iss.group(2)
                    
                    
                    base_de_calculo = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){33}\n.+<ns3:BaseCalculo>\n.+ (.+)').search(arq)
                    if not base_de_calculo:
                        # com nota substituída
                        base_de_calculo = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){41}\n.+<ns3:BaseCalculo>\n.+ (.+)').search(arq)
                        if not base_de_calculo:
                            # com pis cofins
                            base_de_calculo = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){45}\n.+<ns3:BaseCalculo>\n.+ (.+)').search(arq)
                            if not base_de_calculo:
                                # com nota substituída e pis cofins
                                base_de_calculo = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){53}\n.+<ns3:BaseCalculo>\n.+ (.+)').search(arq)
                    
                    base_de_calculo = base_de_calculo.group(2)
                    
                    
                    aliquota = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){36}\n.+<ns3:Aliquota>\n.+ (.+)').search(arq)
                    if not aliquota:
                        # com nota substituída
                        aliquota = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){44}\n.+<ns3:Aliquota>\n.+ (.+)').search(arq)
                        if not aliquota:
                            # com pis cofins
                            aliquota = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){48}\n.+<ns3:Aliquota>\n.+ (.+)').search(arq)
                            if not aliquota:
                                # com nota substituída e pis cofins
                                aliquota = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){56}\n.+<ns3:Aliquota>\n.+ (.+)').search(arq)

                    aliquota = aliquota.group(2)
                    
                    
                    valor_liquido = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){39}\n.+<ns3:ValorLiquidoNfse>\n.+ (.+)').search(arq)
                    if not valor_liquido:
                        # com nota substituída
                        valor_liquido = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){47}\n.+<ns3:ValorLiquidoNfse>\n.+ (.+)').search(arq)
                        if not valor_liquido:
                            # com pis cofins
                            valor_liquido = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){51}\n.+<ns3:ValorLiquidoNfse>\n.+ (.+)').search(arq)
                            if not valor_liquido:
                                # com nota substituída e pis cofins
                                valor_liquido = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){59}\n.+<ns3:ValorLiquidoNfse>\n.+ (.+)').search(arq)
                
                    valor_liquido = valor_liquido.group(2)


                    discriminacao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){52}\n.+<ns3:Discriminacao>\n.+  (.+)').search(arq)
                    if not discriminacao:
                        # com nota substituída
                        discriminacao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){60}\n.+<ns3:Discriminacao>\n.+  (.+)').search(arq)
                        if not discriminacao:
                            # com pis cofins
                            discriminacao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){64}\n.+<ns3:Discriminacao>\n.+  (.+)').search(arq)
                            if not discriminacao:
                                # com nota substituída e pis cofins
                                discriminacao = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){72}\n.+<ns3:Discriminacao>\n.+  (.+)').search(arq)

                    discriminacao = discriminacao.group(2)


                    cnpj_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){61}\n.+<ns3:Cnpj>\n.+ (.+)').search(arq)
                    if not cnpj_prestador:
                        # com nota substituída
                        cnpj_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){69}\n.+<ns3:Cnpj>\n.+ (.+)').search(arq)
                        if not cnpj_prestador:
                            # com pis cofins
                            cnpj_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){73}\n.+<ns3:Cnpj>\n.+ (.+)').search(arq)
                            if not cnpj_prestador:
                                # com nota substituída e pis cofins
                                cnpj_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){81}\n.+<ns3:Cnpj>\n.+ (.+)').search(arq)

                    cnpj_prestador = cnpj_prestador.group(2)


                    nome_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){68}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                    if not nome_prestador:
                        # com nota substituída
                        nome_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){76}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                        if not nome_prestador:
                            # com pis cofins
                            nome_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){80}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                            if not nome_prestador:
                                # com nota substituída e pis cofins
                                nome_prestador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){88}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)

                    nome_prestador = nome_prestador.group(2)
                    

                    cnpj_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){105}\n.+<ns3:CpfCnpj>\n.+\n.+ (.+)').search(arq)
                    if not cnpj_tomador:
                        # com nota substituída
                        cnpj_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){113}\n.+<ns3:CpfCnpj>\n.+\n.+ (.+)').search(arq)
                        if not cnpj_tomador:
                            # com pis cofins
                            cnpj_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){117}\n.+<ns3:CpfCnpj>\n.+\n.+ (.+)').search(arq)
                            if not cnpj_tomador:
                                # com nota substituída e pis cofins
                                cnpj_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){125}\n.+<ns3:CpfCnpj>\n.+\n.+ (.+)').search(arq)

                    cnpj_tomador = cnpj_tomador.group(2)


                    nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){111}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                    if not nome_tomador:
                        # com nota substituída
                        nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){119}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                        if not nome_tomador:
                            # com pis cofins
                            nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){124}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                            if not nome_tomador:
                                # com pis cofins tomador cnpj
                                nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){126}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)
                                if not nome_tomador:
                                    # com nota substituída e pis cofins
                                    nome_tomador = re.compile(r'<ns3:IdentificacaoNfse>\n.+ <ns3:Numero>\n.+ ' + nota + '\n.+ </ns3:Numero>(\n.+){134}\n.+<ns3:RazaoSocial>\n.+  (.+)').search(arq)

                    nome_tomador = nome_tomador.group(2)
                    

                    _escreve_relatorio_csv(f'{nota};{nota_substituida};{data_emissao};{competencia};{cnpj_prestador};{nome_prestador};{cnpj_tomador};{nome_tomador};{discriminacao};'
                                           f'{valor_servico};{valor_liquido};{valor_iss};{base_de_calculo};{aliquota}')
                            
                    
            except():
                print(f'\nArquivo: {arq} - ERRO')
            

                
                
if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    _escreve_header_csv(';'.join(['NOTA', 'NOTA SUBSTITUÍDA', 'EMISSÃO', 'COMPETÊNCIA', 'CNPJ PRESTADOR', 'RAZÃO SOCIAL PRESTADOR', 'CPF / CNPJ TOMADOR',
                                  'RAZÃO SOCIAL TOMADOR', 'DISCRIMINAÇÃO', 'VALOR', 'VALOR LIQUIDO', 'VALOR ISS', 'BASE DE CÁLCULO', 'ALIQUOTA']))
    print(datetime.now() - inicio)
