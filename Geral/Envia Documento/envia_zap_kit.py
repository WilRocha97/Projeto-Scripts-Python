# -*- coding: utf-8 -*-
import ramdom, time, os, shutil, pyperclip, pywhatkit

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers
from chrome_comum import _initialize_chrome


def verifica_o_numero(arquivo):
    # Abrir o pdf
    with fitz.open(arquivo) as pdf:
        # Para cada página do pdf
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                print(textinho)
                cnpj = ''
            except:
                print(textinho)

    # Definir o nome das colunas
    colunas = ['cnpj', 'numero']
    # Localiza a planilha
    caminho = Path('V:/Setor Robô/Scripts Python/Geral/WhatsApp/ignore/Dados.csv')
    # Abrir a planilha
    lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
    # Definir o index da planilha
    lista.set_index('cnpj', inplace=True)

    numero = lista.loc[int(cnpj)]
    numero = list(numero)
    numero = numero[0]
    
    return cnpj, numero


def envia(numero, arquivo):
    try:
        pywhatkit.sendwhatmsg_instantly('+55' + numero, "hello", 12, True, 3)
        return True
    except:
        return False
    

@_time_execution
def run():
    download_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\ignore\\arquivos"
    final_folder_enviado = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\Execução\\Enviados"
    final_folder_nao_enviado = "V:\\Setor Robô\\Scripts Python\\Geral\\WhatsApp\\Execução\\Não Enviados"
    cnpj = 'Facilita'
    while 0 < 1:
        print('>>> Aguardando documentos...')
        # limpa a pasta de download caso fique algum arquivo nela
        for arq in os.listdir(download_folder):
            arquivo = os.path.join(download_folder, arq)
            
            # determina o tempo de espera entre uma mensagem e outra para tentar evitar span
            numero = random.randint(1, 10)
            time.sleep(numero)
            
            # numero = verifica_o_numero(arquivo)
            numero = '19 98436-0240'
        
            print(numero)
            if numero:
                if envia(numero, arquivo):
                    os.makedirs(final_folder_enviado, exist_ok=True)
                    time.sleep(1)
                    shutil.move(arquivo, os.path.join(final_folder_enviado, arq))
                    _escreve_relatorio_csv(f'{cnpj};{numero};Arquivo enviado', nome='Envia guia')
                else:
                    os.makedirs(final_folder_nao_enviado, exist_ok=True)
                    time.sleep(1)
                    shutil.move(arquivo, os.path.join(final_folder_nao_enviado, arq))
                    _escreve_relatorio_csv(f'{cnpj};{numero};Erro ao enviar o arquivo', nome='Envia guia')
            else:
                os.makedirs(final_folder_nao_enviado, exist_ok=True)
                time.sleep(1)
                shutil.move(arquivo, os.path.join(final_folder_nao_enviado, arq))
                _escreve_relatorio_csv(f'{cnpj};{numero};Número não encontrado', nome='Envia guia')
    

if __name__ == '__main__':
    run()

