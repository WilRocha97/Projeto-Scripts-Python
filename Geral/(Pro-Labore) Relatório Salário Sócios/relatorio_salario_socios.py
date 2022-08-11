# -*- coding: utf-8 -*-
import fitz, re, os
from datetime import datetime

# Definir os padrões de regex
padrao_nome = re.compile(r'(\d{1,5})(.+)\nCód\. Sócio')

# Caminho dos arquivos que serão analizados
arquivos = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', '(Pro-Labore) Relatório Salário Sócios', 'PDF')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Salário Sócios.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Salário Sócios - Auxiliar.csv', 'a') as f:
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

                    # Procura o nome da empresa
                    localiza_nome = padrao_nome.search(textinho)

                    if not localiza_nome:
                        empresa = 'Sem Nome'
                        cod = 'Página: ' + str(page.number)
                    else:
                        # Guarda as infos da empresa
                        empresa = localiza_nome.group(2)
                        cod = localiza_nome.group(1)

                    l_nome_socio = re.findall(r'\n\d{1,4}\n(.+)', textinho)
                    nome_anterior = ''
                    for nome in l_nome_socio:
                        if nome != 'DPCUCA 2022 A - www.cucafresca.com.br / VEIGA & POSTAL (019)3829-8959':
                            if nome != nome_anterior:

                                try:
                                    print('{} - {} - {}\n'.format(cod, empresa, nome))
                                    texto = ';'.join([cod, empresa, '0000', nome])
                                    escrever(texto)
                                except:
                                    pass

                                try:
                                    l_data_nome = re.findall(r'(\d{2}\/\d{2}\/\d{4})\n(\d{1,4})\n(' + nome + ')', textinho)
                                    for info in l_data_nome:
                                        data = info[0]
                                        codigo_socio = info[1]

                                        print('{} - {} - {} - {} - {} - {}\n'.format(cod, empresa, codigo_socio, nome, data, 'Salário Mínimo'))
                                        texto = ';'.join([cod, empresa, codigo_socio, nome, data, 'Salário Mínimo'])
                                        escrever(texto)
                                except:
                                    pass

                                try:
                                    l_data_valor_nome = re.findall(r'(.+)\n +(.+)\n(\d{1,4})\n(' + nome + ')', textinho)

                                    for info in l_data_valor_nome:
                                        data = info[0]
                                        valor = info[1]
                                        codigo_socio = info[2]

                                        print('{} - {} - {} - {} - {} - R${}\n'.format(cod, empresa, codigo_socio, nome, data, valor))
                                        texto = ';'.join([cod, empresa, codigo_socio, nome, data, valor])
                                        escrever(texto)
                                except:
                                    pass

                                nome_anterior = nome
                except():
                    print('ERRO---------------------------------\n' + textinho)


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
