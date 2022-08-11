# -*- coding: utf-8 -*-
import fitz, re, os
from datetime import datetime

# Definir os padrões de regex
padrao_nome = re.compile(r'CNPJ:\n.+\n(.+)\n.+\n(.+)')
p_pagamento_pro_labore = re.compile(r'(Pagamento Pro Labore)\n.+\n +(.+)')
p_irf_pro_labore = re.compile(r'(IRF s/ pro labore)\n.+\n +(.+)')
p_inss_segurado_individual = re.compile(r'(INSS Segurado Contrib.Individual)\n.+\n +(.+)')
p_inss_empresa = re.compile(r'(INSS Empresa)\n.+\n +(.+)')
p_inss_empresa_base = re.compile(r'(INSS Empresa- Base:) +(.+)')

# Caminho dos arquivos que serão analizados
arquivos = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', '(Pro-Labore) Resumo de pagamento Sócios', 'PDF')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Resumo Pagamento Socios.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Resumo Pagamento Socios - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')


def analiza():
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(arquivos):
        print('\nArquivo: {} \n'.format(arq))
        # Abrir o pdf
        with fitz.open(r'PDF\{}'.format(arq)) as pdf:

            # Pra cada pagina do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.getText('text', flags=1 + 2 + 8)
                    # print('---------------------------------------------------------------------------------------\n\n\n' + textinho)
                    # Procura o valor a recolher da empresa
                    localiza_nome_socio = re.findall(r'(.+)\nLocal', textinho)

                    if not localiza_nome_socio:
                        continue

                    # Procura o nome da empresa
                    localiza_nome = padrao_nome.search(textinho)

                    # Guarda as infos da empresa
                    empresa = localiza_nome.group(2)
                    cnpj = localiza_nome.group(1)
                    for count, nome in enumerate(localiza_nome_socio):
                        nome_socio = nome

                        try:
                            l_pagamento_pro_labore = re.findall(p_pagamento_pro_labore, textinho)

                            descricao = l_pagamento_pro_labore[count][0]
                            valor = l_pagamento_pro_labore[count][1]

                            print('{} - {} - {} - {} - R${}\n'.format(empresa, cnpj, nome_socio, descricao, valor))
                            texto = ';'.join([empresa, cnpj, nome_socio, descricao, valor])
                            escrever(texto)
                        except:
                            pass

                        try:
                            l_irf_pro_labore = re.findall(p_irf_pro_labore, textinho)

                            descricao = l_irf_pro_labore[count][0]
                            valor = l_irf_pro_labore[count][1]

                            print('{} - {} - {} - {} - R${}\n'.format(empresa, cnpj, nome_socio, descricao, valor))
                            texto = ';'.join([empresa, cnpj, nome_socio, descricao, valor])
                            escrever(texto)
                        except:
                            pass

                        try:
                            l_inss_empresa = re.findall(p_inss_empresa, textinho)

                            descricao = l_inss_empresa[count][0]
                            valor = l_inss_empresa[count][1]

                            print('{} - {} - {} - {} - R${}\n'.format(empresa, cnpj, nome_socio, descricao, valor))
                            texto = ';'.join([empresa, cnpj, nome_socio, descricao, valor])
                            escrever(texto)
                        except:
                            pass

                        try:
                            l_inss_empresa_base = re.findall(p_inss_empresa_base, textinho)

                            descricao = l_inss_empresa_base[count][0]
                            valor = l_inss_empresa_base[count][1]

                            print('{} - {} - {} - {} - R${}\n'.format(empresa, cnpj, nome_socio, descricao, valor))
                            texto = ';'.join([empresa, cnpj, nome_socio, descricao, valor])
                            escrever(texto)
                        except:
                            pass

                        try:
                            l_inss_segurado_individual = re.findall(p_inss_segurado_individual, textinho)

                            descricao = l_inss_segurado_individual[count][0]
                            valor = l_inss_segurado_individual[count][1]

                            print('{} - {} - {} - {} - R${}\n'.format(empresa, cnpj, nome_socio, descricao, valor))
                            texto = ';'.join([empresa, cnpj, nome_socio, descricao, valor])
                            escrever(texto)
                        except:
                            pass

                except():
                    print('ERRO---------------------------------\n' + textinho)


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
