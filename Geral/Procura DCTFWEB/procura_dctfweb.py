# -*- coding: utf-8 -*-
import fitz, re, os
import shutil
from datetime import datetime

# Definir os padrões de regex
padrao_pendencia = re.compile(r'(.+) - (CP-SEGUR.)\n(.+)\n.+\n(.+)\n(.+)')
padrao_pendencia2 = re.compile(r'(.+) - (CP-PATRONAL)\n(.+)\n.+\n(.+)\n(.+)')
padrao_cnpj = re.compile(r'CNPJ: (\d\d.\d\d\d.\d\d\d) - ')

# Caminho dos arquivos que serão analizados
#arquivos_origem = os.path.join(r'\\VPSRV02', 'Comum', 'RELATÓRIO DE PESQUISAS', 'RELATÓRIO RFB 25.11.2021', 'DEZEMBRO-2021', 'PENDENCIAS')
#arquivos_destino = os.path.join(r'\\VPSRV02', 'Comum', 'RELATÓRIO DE PESQUISAS', 'RELATÓRIO RFB 25.11.2021', 'DEZEMBRO-2021', 'PENDENCIAS DCTF WEB')
arquivos_origem = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', 'Procura DCTFWEB', 'PENDENCIAS')
arquivos_destino = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Scripts Python', 'Geral', 'Procura DCTFWEB', 'PENDENCIAS DCTF WEB')


def escrever(texto):
    # Anota o andamento na planilha
    try:
        with open('Pendencias DCTF WEB.csv', 'a') as f:
            f.write(texto + '\n')
    except():
        with open('Pendencias DCTF WEB - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')


def localiza(arq, cnpj, localiza_pendencia):
    for loc in localiza_pendencia:
        receita = loc[0]
        pendencia = loc[1]
        comp = loc[2]
        valor_original = loc[3]
        saldo_devedor = loc[4]

        texto = ';'.join([arq, str(cnpj), str(receita), str(pendencia), str(comp), str(valor_original), str(saldo_devedor)])
        escrever(texto)


def analiza():
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(arquivos_origem):
        print('\nArquivo: {}'.format(arq))
        # Abrir o pdf
        with fitz.open(r'{}\{}'.format(arquivos_origem, arq)) as pdf:
            x = 0
            # Pra cada pagina do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.getText('text', flags=1 + 2 + 8)
                    # Procura o valor a recolher da empresa
                    pega_cnpj = padrao_cnpj.search(textinho)
                    cnpj = pega_cnpj.group(1)

                    localiza_pendencia = re.findall(padrao_pendencia, textinho)
                    if localiza_pendencia:
                        x += 1
                        localiza(arq, cnpj, localiza_pendencia)

                    localiza_pendencia2 = re.findall(padrao_pendencia2, textinho)
                    if localiza_pendencia2:
                        x += 1
                        localiza(arq, cnpj, localiza_pendencia2)

                except():
                    print('ERRO')
                    print(textinho)

        arquivo_origem = os.path.join(arquivos_origem, arq)
        arquivo_destino = os.path.join(arquivos_destino, arq)
        if x >= 1:
            shutil.move(arquivo_origem, arquivo_destino)
            print('✔ - Termo encontrado - {}\n'.format(arq))
        else:
            print('❌ - Não encontrado - {}\n'.format(arq))


if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    print(datetime.now() - inicio)
