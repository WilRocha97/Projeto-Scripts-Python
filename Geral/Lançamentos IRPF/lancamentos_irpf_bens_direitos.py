# -*- coding: utf-8 -*-
import fitz, atexit, sys, time, re, os, shutil, warnings, pyarrow, pyautogui as p, pandas as pd, PySimpleGUI as sg
from datetime import datetime
from threading import Thread
from functools import wraps
from unidecode import unidecode
from random import choice
from PIL import Image
from win32com import client
from pathlib import Path

e_dir = Path('execução')


def create_lock_file(lock_file_path):
    try:
        # Tente criar o arquivo de trava
        with open(lock_file_path, 'x') as lock_file:
            lock_file.write(str(os.getpid()))
        return True
    except FileExistsError:
        # O arquivo de trava já existe, indicando que outra instância está em execução
        return False


def remove_lock_file(lock_file_path):
    try:
        os.remove(lock_file_path)
    except FileNotFoundError:
        pass


def open_lista_dados(file, encode='latin-1'):
    if not file:
        return False

    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


# Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'
def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


def indice(count, total_empresas, empresa, index=0):
    if count > 1:
        print(f'[ {len(total_empresas) - (count - 1)} Restantes ]\n\n')
    # Cria um indice para saber qual linha dos dados está
    indice_dados = f'[ {str(count + index)} de {str(len(total_empresas) + index)} ]'
    
    empresa = str(empresa).replace("('", '[ ').replace("')", ' ]').replace("',)", " ]").replace(',)', ' ]').replace("', '", ' - ')
    
    print(f'{indice_dados} - {empresa}')


def concatena(variavel, quantidade, posicao, caractere):
    variavel = str(variavel)
    if posicao == 'depois':
        # concatena depois
        while len(str(variavel)) < quantidade:
            variavel += str(caractere)
    if posicao == 'antes':
        # concatena antes
        while len(str(variavel)) < quantidade:
            variavel = str(caractere) + str(variavel)
            
    return variavel


def numero_controle_existe(numero_pesquisado):
    try:
        # Definir o nome das colunas
        colunas = ['numero']
        # Localiza a planilha
        caminho = Path('ignore/numeros_controle.csv')
        # Abrir a planilha
        lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
        # Definir o index da planilha
        lista.set_index('numero', inplace=True)
        numero = lista.loc[int(numero_pesquisado)]

        return numero
    except:
        return False


def captura_chave_item_existentes(arquivo_original, arquivo_final):
    try:
        with open(arquivo_final, 'r') as arquivo:
            # Leia o conteúdo do arquivo
            conteudo = arquivo.readlines()
    except:
        try:
            with open(arquivo_original, 'r') as arquivo:
                # Leia o conteúdo do arquivo
                conteudo = arquivo.readlines()
        except:
            return False, False
            
    # verifica os números de chave que podem já existir no arquivo original
    lista_chaves = []
    for linha in conteudo:
        if str(linha[:2]) == '27':
            lista_chaves.append(int(linha[891:896]))
    
    # retorna o conteudo do arquivo original e o maior número encontrado na lista de chaves
    return conteudo, max(int(number) for number in lista_chaves) + 1

  
def busca_codigo_municipio(municipio):
    try:
        # Definir o nome das colunas
        colunas = ['município', 'uf', 'código', 'nome']
        # Localiza a planilha
        caminho = Path('V:\Setor Robô\Scripts Python\Geral\Lançamentos IRPF\Municípios.csv')
        # Abrir a planilha
        lista = pd.read_csv(caminho, header=None, names=colunas, sep=';', encoding='latin-1')
        # Definir o index da planilha
        lista.set_index('município', inplace=True)
        infos = lista.loc[str(municipio)]
        
        infos = list(infos)
        uf = infos[0]
        codigo = infos[1]
        nome = infos[2]
        
        return [uf, codigo, nome]
    except:
        return False
  
    
def cria_bloco_de_informacoes_b3(arquivo_original, arquivo_final, final_folder, cpf, ano_base, ano_anterior):
    print('>>> Analisando planilhas')
    novas_linhas = ''
    
    # função para capturar as chaves dos itens no arquivo 'DBK' e gerar uma nova chave inicial original e para capturar o conteudo original
    conteudo_arquivo_original, chave_item = captura_chave_item_existentes(arquivo_original, arquivo_final)
    if not conteudo_arquivo_original:
        return f'Arquivo referente ao CPF {cpf} não encontrado'
    
    
    # as planilhas possuem duas abas cada, deve percorrer as duas
    for aba in ['Posição - Ações', 'Posição - Fundos']:
        # inicia as variáveis para não dar alerta
        endereco = ''
        endereco_2 = ''
        imovel = ''
        nome_cartorio = ''
        valor_ano_principal = ''
        valor_ano_anterior = ''
        cnpj_empresa = ''
        veiculos = ''
        espaco = ''
        codigo_tipo = ''
        grupo_tipo = ''
        
        # para cada planilha na pasta cria uma lista com elas
        try:
            planilha_ano_anterior = pd.read_excel(os.path.join(final_folder, f'{str(cpf)} - relatorio-consolidado-anual-{ano_anterior}.xlsx'), sheet_name=aba)
        except:
            return f'Planilha "{str(cpf)} - relatorio-consolidado-anual-{ano_anterior}.xlsx" não encontrada'
        try:
            planilha_principal = pd.read_excel(os.path.join(final_folder, f'{str(cpf)} - relatorio-consolidado-anual-{str(ano_base)}.xlsx'), sheet_name=aba)
        except:
            return f'Planilha "{str(cpf)} - relatorio-consolidado-anual-{ano_base}.xlsx" não encontrada'
        # percorre a planilha principal e para cada linha procura um item igual na planilha secundaria, se encontrar anota o item com os valores das duas planilhas
        # marca a variável 'semelhante' como 'sim' para indicar que aquela linha tem semelhante, se sair do for da segunda planilha com a variável 'semelhante' marcada
        # como 'não', anota apenas as infos da linha da planilha principal e o valor da planilha secundária fica 0000000000000
        for key_1 in planilha_principal.itertuples():
            semelhante = 'não'
            
            # concatena essa variável para criar um número único para a chave do item
            chave_item = int(chave_item)
            chave_item += 1
            
            # se a célula da planilha estiver vazia, pula para a próxima linha
            if str(key_1[5]) == 'nan':
                continue
            # percorre a segunda planilha para procurar o semelhante
            for key_2 in planilha_ano_anterior.itertuples():
                if str(key_2[5]) == 'nan':
                    continue
                    
                cnpj_empresa = str(key_1[5]).replace('.0', '')
                
                # Converta o número float em uma string formatada
                valor_ano_principal = "{:.2f}".format(key_1[14])
                valor_ano_anterior = "{:.2f}".format(key_2[14])
                
                # formata os valores retirando pontuação, pois no arquivo 'DBK' não pode ter
                valor_ano_principal = valor_ano_principal.replace('.', '')
                valor_ano_anterior = valor_ano_anterior.replace('.', '')

                # ações 01 - fundos 03 código para colocar depois da primeira ocorrencia do CPF do titular
                # ações fundos 02 ou ações cotas 03 - fundos 07 código para colocar depois da segunda ocorrencia do CPF do titular
                if aba == 'Posição - Ações':
                    codigo_tipo = '01'  # depois da primeira ocorrencia do CPF
                    grupo_tipo = '03'  # depois da segunda ocorrencia do CPF
                
                if aba == 'Posição - Fundos':
                    if key_1[7] == 'Fundo':
                        codigo_tipo = '02' # depois da primeira ocorrencia do CPF
                    if key_1[7] == 'Cotas':
                        codigo_tipo = '03' # depois da primeira ocorrencia do CPF
                    grupo_tipo = '07'  # depois da segunda ocorrencia do CPF
                    
                # se encontrar um semelhante anota os valores dos dois
                if key_1[5] == key_2[5]:
                    if key_1[2] == key_2[2]:
                        descricao_bem = str(key_1[1])
                        # formata a informação para remover caracteres acentuados, pois o programa do IRPF não aceita arquivos com texto acentuado
                        descricao_bem = unidecode(descricao_bem)
                        
                        # na criação da linha, chama a função 'concatena()' para concatenar espaços em branco para que a linha fique formatada corretamente
                        nova_linha = (f'27{cpf}{codigo_tipo}0105'
                                f'{concatena(descricao_bem, 512, "depois", " ")}'
                                f'{concatena(valor_ano_anterior, 13, "antes", "0")}'
                                f'{concatena(valor_ano_principal, 13, "antes", "0")}'
                                f'{concatena(endereco, 137, "depois", " ")}0000'
                                f'{concatena(endereco_2, 40, "depois", " ")}2'
                                f'{concatena(imovel, 80, "depois", " ")}000000000002'
                                f'{concatena(nome_cartorio, 60, "depois", " ")}'
                                f'{concatena(chave_item, 5, "antes", "0")}00000000'
                                f'{concatena(veiculos, 118, "depois", " ")}0000'
                                f'{concatena(espaco, 15, "depois", " ")}'
                                f'{concatena(cnpj_empresa, 14, "antes", "0")}'
                                f'{concatena(espaco, 30, "depois", " ")}000T{cpf}{grupo_tipo}0'
                                f'{concatena(espaco, 28, "depois", " ")}0000000000001'
                                f'{concatena(key_1[4], 20, "depois", " ")}'
                                f'{concatena(espaco, 10, "depois", " ")}\n')
                        # concatena a linha nova dentro da váriavel para depois passar todas elas direto para a função de aditar o arquivo original 'DBK'
                        novas_linhas += nova_linha
                        
                        # anota que encontrou semelhante para não entrar no if abaixo e sai da repetição
                        semelhante = 'sim'
                        break
                    
            if semelhante == 'não':
                descricao_bem = str(key_1[1])
                descricao_bem = unidecode(descricao_bem)
                
                # na criação da linha, chama a função 'concatena()' para concatenar espaços em branco para que a linha fique formatada corretamente
                nova_linha = (f'27{cpf}{codigo_tipo}0105'
                        f'{concatena(descricao_bem, 512, "depois", " ")}'
                        f'0000000000000'
                        f'{concatena(valor_ano_principal, 13, "antes", "0")}'
                        f'{concatena(endereco, 137, "depois", " ")}0000'
                        f'{concatena(endereco_2, 40, "depois", " ")}2'
                        f'{concatena(imovel, 80, "depois", " ")}000000000002'
                        f'{concatena(nome_cartorio, 60, "depois", " ")}'
                        f'{concatena(chave_item, 5, "antes", "0")}00000000'
                        f'{concatena(veiculos, 118, "depois", " ")}0000'
                        f'{concatena(espaco, 15, "depois", " ")}'
                        f'{concatena(cnpj_empresa, 14, "antes", "0")}'
                        f'{concatena(espaco, 30, "depois", " ")}000T{cpf}{grupo_tipo}0'
                        f'{concatena(espaco, 28, "depois", " ")}0000000000001'
                        f'{concatena(key_1[4], 20, "depois", " ")}'
                        f'{concatena(espaco, 10, "depois", " ")}\n')
                # concatena a linha nova dentro da váriavel para depois passar todas elas direto para a função de aditar o arquivo original 'DBK'
                novas_linhas += nova_linha
            
        # faz e mesma lógica, porém ao invés de escrever no arquivo quando encontrar o semelhante, apenas os diferêntes serão escritos
        for key_2 in planilha_ano_anterior.itertuples():
            semelhante = 'não'
            
            # concatena essa variável para criar um número único para a chave do item
            chave_item = int(chave_item)
            chave_item += 1
            
            # se a célula da planilha estiver vazia, pula para a próxima linha
            if str(key_2[5]) == 'nan':
                continue
            
            # percorre a primeira planilha para procurar o semelhante
            for key_1 in planilha_principal.itertuples():
                if str(key_1[5]) == 'nan':
                    continue
                
                cnpj_empresa = str(key_1[5]).replace('.0', '')
                
                # Converta o número float em uma string formatada
                valor_ano_anterior = "{:.2f}".format(key_2[14])
                
                # formata os valores retirando pontuação, pois no arquivo 'DBK' não pode ter
                valor_ano_anterior = valor_ano_anterior.replace('.', '')
                
                # ações 01 - fundos 03 código para colocar depois da primeira ocorrencia do CPF do titular
                # ações fundos 02 ou ações cotas 03 - fundos 07 código para colocar depois da segunda ocorrencia do CPF do titular
                if aba == 'Posição - Ações':
                    codigo_tipo = '01'  # depois da primeira ocorrencia do CPF
                    grupo_tipo = '03'  # depois da segunda ocorrencia do CPF
                    
                if aba == 'Posição - Fundos':
                    if key_2[7] == 'Fundo':
                        codigo_tipo = '02'  # depois da primeira ocorrencia do CPF
                    if key_2[7] == 'Cotas':
                        codigo_tipo = '03'  # depois da primeira ocorrencia do CPF
                    grupo_tipo = '07'  # depois da segunda ocorrencia do CPF
                 
                # nesse caso se for semelhante não anota na planilha, pois já foi inserida quando percorreu a primeira planilha
                if key_2[5] == key_1[5]:
                    if key_2[2] == key_1[2]:
                        semelhante = 'sim'
                        break
            
            # anota somente as diferentes da segunda planilha, pois não foram anotadas quando percorreu a primeira planilha
            if semelhante == 'não':
                descricao_bem = str(key_2[1])
                descricao_bem = unidecode(descricao_bem)
                
                # na criação da linha, chama a função 'concatena()' para concatenar espaços em branco para que a linha fique formatada corretamente
                nova_linha = (f'27{cpf}{codigo_tipo}0105'
                        f'{concatena(descricao_bem, 512, "depois", " ")}'
                        f'{concatena(valor_ano_anterior, 13, "antes", "0")}'
                        f'0000000000000'
                        f'{concatena(endereco, 137, "depois", " ")}0000'
                        f'{concatena(endereco_2, 40, "depois", " ")}2'
                        f'{concatena(imovel, 80, "depois", " ")}000000000002'
                        f'{concatena(nome_cartorio, 60, "depois", " ")}'
                        f'{concatena(chave_item, 5, "antes", "0")}00000000'
                        f'{concatena(veiculos, 118, "depois", " ")}0000'
                        f'{concatena(espaco, 15, "depois", " ")}'
                        f'{concatena(cnpj_empresa, 14, "antes", "0")}'
                        f'{concatena(espaco, 30, "depois", " ")}000T{cpf}{grupo_tipo}0'
                        f'{concatena(espaco, 28, "depois", " ")}0000000000001'
                        f'{concatena(key_2[4], 20, "depois", " ")}'
                        f'{concatena(espaco, 10, "depois", " ")}\n')
                # concatena a linha nova dentro da váriavel para depois passar todas elas direto para a função de aditar o arquivo original 'DBK'
                novas_linhas += nova_linha
            
    resultado = insere_nova_linha(arquivo_final, conteudo_arquivo_original, novas_linhas)
    
    return resultado


def cria_bloco_de_informacoes_veiculos_e_imoveis(ano_base, ano_anterior, caminho_planilha, caminho_arquivo_original, caminho_arquivo_editados):
    novas_linhas = ''
    abas_encontradas = []
    aba_imovel = ''
    aba_carro = ''
    
    for aba in ['Imóvel IRPF', 'Carro IRPF']:
        try:
            df = pd.read_excel(caminho_planilha, sheet_name=aba)
        except:
            continue
        
        print(f'>>> Analisando planilha {aba}')
        for linha in df.itertuples():
            cpf = linha[1]
            print(cpf)
            arquivo_original = os.path.join(caminho_arquivo_original, f'{cpf}-IRPF-A-{ano_base}-{ano_anterior}-ORIGI.DBK')
            os.makedirs(os.path.join(caminho_arquivo_editados, 'Arquivos'), exist_ok=True)
            arquivo_final = os.path.join(caminho_arquivo_editados, 'Arquivos', f'{cpf}-IRPF-A-{ano_base}-{ano_anterior}-ORIGI.DBK')
            
            # função para capturar as chaves dos itens no arquivo 'DBK' e gerar uma nova chave inicial original e para capturar o conteudo original
            conteudo_arquivo_original, chave_item = captura_chave_item_existentes(arquivo_original, arquivo_final)
            if not conteudo_arquivo_original:
                return cpf, f'Arquivo referente ao CPF {cpf} não encontrado'
            
            if aba == 'Imóvel IRPF':
                aba_imovel = 'Imóvel IRPF'
                
                codigo_municipio = busca_codigo_municipio(linha[13])
                if not codigo_municipio:
                    print('❌ Código do município não encontrado')
                    return cpf, 'Código do município não encontrado'
                
                registro_do_imovel = ''
                if linha[17] == 'Não':
                    registro_do_imovel = '0'
                if linha[17] == 'Sim':
                    registro_do_imovel = '1'
                if linha[17] == 'Vazio':
                    registro_do_imovel = '2'
                
                if re.compile(r',').search(str(linha[15])):
                    area_imovel = str(linha[15]).replace(',', '')
                else:
                    area_imovel = str(linha[15]) + '0'
                
                unidade_medida = ''
                if linha[16] == 'M²':
                    unidade_medida = '0'
                if linha[16] == 'há':
                    unidade_medida = '1'
                if linha[16] == '':
                    unidade_medida = '2'
            
                data_aquisicao = str(linha[6]).replace(' 00:00:00', '').split('-')
                data_aquisicao = f'{data_aquisicao[2]}{data_aquisicao[1]}{data_aquisicao[0]}'
                
                # Converta o número float em uma string formatada
                valor_ano_anterior = "{:.2f}".format(linha[20])
                # formata os valores retirando pontuação, pois no arquivo 'DBK' não pode ter
                valor_ano_anterior = valor_ano_anterior.replace('.', '')
                # Converta o número float em uma string formatada
                valor_ano_principal = "{:.2f}".format(linha[21])
                # formata os valores retirando pontuação, pois no arquivo 'DBK' não pode ter
                valor_ano_principal = valor_ano_principal.replace('.', '')
            
                # na criação da linha, chama a função 'concatena()' para concatenar espaços em branco para que a linha fique formatada corretamente
                nova_linha = (f'27{cpf}{concatena(linha[3], 2, "antes", "0")}0105'              # CPF e código do lançamento
                              f'{concatena(linha[7], 512, "depois", " ")}'                      # descrição
                              f'{concatena(valor_ano_anterior, 13, "antes", "0")}'              # valor anterior
                              f'{concatena(valor_ano_principal, 13, "antes", "0")}'             # valor atual
                              f'{concatena(linha[8], 40, "depois", " ")}'                       # endereço
                              f'{concatena(linha[9], 6, "depois", " ")}'                        # número
                              f'{concatena(linha[10], 40, "depois", " ")}'                      # complemento
                              f'{concatena(linha[11], 40, "depois", " ")}'                      # bairro
                              f'{concatena(linha[14], 9, "depois", " ")}'                       # CEP
                              f'{codigo_municipio[0]}'                                                                         # UF
                              f'{concatena(codigo_municipio[1], 4, "antes", "0")}'              # código do município
                              f'{concatena(codigo_municipio[2], 40, "depois", " ")}'            # município
                              f'{registro_do_imovel}'                                                                          # imóvel registrado
                              f'{concatena(linha[18], 40, "depois", " ")}'                      # matrícula
                              f'{concatena(" ", 40, "depois", " ")}'                  # filler, espaços vazios
                              f'{concatena(area_imovel, 11, "antes", "0")}'                     # área de imóvel
                              f'{unidade_medida}'                                                                              # unidade de médida
                              f'{concatena(linha[19], 60, "depois", " ")}'                      # nome do cartório
                              f'{concatena(chave_item, 5, "antes", "0")}'                       # chave de identificação do bem
                              f'{concatena(data_aquisicao, 8, "antes", "0")}'                   # data da aquisição
                              f'{concatena("0000", 122, "antes", " ")}'               # filler, espaços vazios e a agência
                              f'{concatena(" ", 29, "depois", " ")}'                  # filler, espaços vazios
                              f'{concatena(linha[5], 30, "depois", " ")}000'                    # IPTU e número do banco
                              f'{concatena(" ", 12, "depois", " ")}'                  # filler, espaços vazios
                              f'{concatena(linha[2], 2, "antes", "0")}0'                        # grupo do bem e 0 se não for inventariar e 1 se sim
                              f'{concatena(" ", 28, "depois", " ")}0000000000000'     # filler, espaços vazios
                              f'{concatena(" ", 20, "depois", " ")}'                  # código de negociação da bolsa
                              f'{concatena(" ", 10, "depois", " ")}\n')               # número de controle, o sistema aceita se estiver vazio
            
            if aba == 'Carro IRPF':
                aba_carro = 'Carro IRPF'
                
                # Converta o número float em uma string formatada
                valor_ano_anterior = "{:.2f}".format(linha[7])
                # formata os valores retirando pontuação, pois no arquivo 'DBK' não pode ter
                valor_ano_anterior = valor_ano_anterior.replace('.', '')
                
                # na criação da linha, chama a função 'concatena()' para concatenar espaços em branco para que a linha fique formatada corretamente
                nova_linha = (f'27{cpf}{concatena(linha[3], 2, "antes", "0")}0105'              # CPF e código do lançamento
                              f'{concatena(linha[6], 512, "depois", " ")}'                      # descrição
                              f'{concatena(valor_ano_anterior, 13, "antes", "0")}'              # valor anterior
                              f'{concatena("0", 13, "antes", "0")}'                   # valor atual
                              f'{concatena(" ", 137, "antes", " ")}0000'              # filler, espaços vazios e código do município
                              f'{concatena(" ", 40, "antes", " ")}2'                  # filler, espaços vazios e indicativo de registro de imovel
                              f'{concatena(" ", 80, "antes", " ")}000000000002'       # filler, espaços vazios, aréa do imóvel e unidade de medida
                              f'{concatena(" ", 60, "antes", " ")}'                   # filler, espaços vazios
                              f'{concatena(chave_item, 5, "antes", "0")}00000000'               # chave de identificação do bem e data de aquisição
                              f'{concatena(" ", 28, "antes", " ")}'                   # filler, espaços vazios
                              f'{concatena(linha[5], 30, "depois", " ")}'                       # renavam
                              f'{concatena(" ", 60, "antes", " ")}0000'               # filler, espaços vazios e número da agência
                              f'{concatena(" ", 59, "antes", " ")}000'                # filler, espaços vazios, número do banco e identificador do registro
                              f'{concatena(" ", 12, "antes", " ")}'                   # filler, espaços vazios
                              f'{concatena(linha[2], 2, "antes", "0")}0'                        # grupo do item
                              f'{concatena(" ", 28, "antes", " ")}'                   # filler, espaços vazios
                              f'{concatena("0", 13, "depois", "0")}'                  # informações da bolsa
                              f'{concatena(" ", 20, "depois", " ")}'                  # código de negociação da bolsa
                              f'{concatena(" ", 10, "depois", " ")}\n')               # número de controle, o sistema aceita se estiver vazio
                
            # concatena a linha nova dentro da váriavel para depois passar todas elas direto para a função de aditar o arquivo original 'DBK'
            novas_linhas += nova_linha
    
    try:
        resultado = insere_nova_linha(arquivo_final, conteudo_arquivo_original, novas_linhas)
    except:
        print('❗ Planilha aberta detectada. Isso não interfere no processo do robô, a não ser que você não tenha salvo as alterações.')
        return False, ''
    
    return cpf, resultado + f';{aba_imovel}' + f';{aba_carro}'


def insere_nova_linha(arquivo_final, conteudo_arquivo_original, nova_linha):
    print('>>> Adicionando informações no arquivo .DBK')
    parte_anterior = ''
    parte_posterior = ''
    # Divida o conteúdo em duas partes: antes e depois do ponto de inserção
    segunda_parte = 'não'
    for linha in conteudo_arquivo_original:
        if segunda_parte == 'não':
            try:
                if int(linha[:2]) <=  27:
                    parte_anterior += linha
                else:
                    segunda_parte = 'sim'
            except:
                parte_anterior += linha
        else:
            parte_posterior += linha
    
    # Insira as novas linhas
    conteudo_modificado = parte_anterior + nova_linha + parte_posterior
    
    # Abra o arquivo para escrita
    with open(arquivo_final, 'w') as arquivo:
        # Escreva o conteúdo modificado de volta no arquivo
        arquivo.writelines(conteudo_modificado)
    
    print('✔ Arquivo editado com sucesso\n')
    return 'Cópia de segurança ".DBK" pronta para importação, '


def run_acoes_fundos(window_acoes_fundos, event, dados_b3, final_folder, caminho_arquivo_original, caminho_arquivo_editados):
    ano_atual = datetime.now().year
    ano_base = ano_atual
    ano_anterior = ano_atual - 1

    andamentos = f'Lançamentos Ações-Fundos IRPF {ano_atual}'

    # abrir a planilha de dados
    empresas = open_lista_dados(dados_b3)
    if not empresas:
        return False
    
    total_empresas = empresas[0:]
    for count, empresa in enumerate(empresas[0:], start=1):
        # configurar o indice para localizar em qual empresa está
        indice(count, total_empresas, empresa, 0)
        cpf, nome = empresa
        
        # Abra o arquivo para leitura
        arquivo_original = os.path.join(caminho_arquivo_original, f'{cpf}-IRPF-A-{ano_base}-{ano_anterior}-ORIGI.DBK')
        os.makedirs(os.path.join(caminho_arquivo_editados, 'Arquivos'), exist_ok=True)
        arquivo_final = os.path.join(caminho_arquivo_editados, 'Arquivos', f'{cpf}-IRPF-A-{ano_base}-{ano_anterior}-ORIGI.DBK')

        resultado = cria_bloco_de_informacoes_b3(arquivo_original, arquivo_final, final_folder, cpf, ano_base - 1, ano_anterior - 1)
        resultado_planilhas = f'{cpf};Lançamentos com planilhas já baixadas;{ano_base - 1},{ano_anterior - 1}'
            
        escreve_relatorio_csv(f'{resultado_planilhas};{resultado}', nome=andamentos, local=caminho_arquivo_editados)
        
        window_acoes_fundos['-progressbar-'].update_bar(count, max=int(len(empresas[0:])))
        window_acoes_fundos['-Progresso_texto-'].update(str(round(float(count) / int(len(empresas[0:])) * 100, 1)) + '%')
        window_acoes_fundos.refresh()
        
        if event == '-Encerrar-':
            return
        
        
def run_veiculos_moveis(window_veiculos_imoveis, event, pasta_planilhas_veiculos_imoveis, caminho_arquivo_original, caminho_arquivo_editados):
    ano_atual = datetime.now().year
    ano_base = ano_atual
    ano_anterior = ano_atual - 1
    
    andamentos = f'Lançamentos Imóveis e Carros IRPF {ano_atual}'
    pasta_das_planilhas = os.listdir(pasta_planilhas_veiculos_imoveis)
    
    for count , planilha in enumerate(pasta_das_planilhas):
        caminho_planilha = os.path.join(pasta_planilhas_veiculos_imoveis, planilha)
        cpf, resultado = cria_bloco_de_informacoes_veiculos_e_imoveis(ano_base, ano_anterior, caminho_planilha, caminho_arquivo_original, caminho_arquivo_editados)
        # verifica se o CPF for false, caso seja, significa que existe alguma planilha aberta e o robô passou pleo backup do save que o windows cria
        if not cpf:
            continue
            
        escreve_relatorio_csv(f'{cpf};{resultado}', nome=andamentos, local=caminho_arquivo_editados)
        
        window_veiculos_imoveis['-progressbar-'].update_bar( count, max=int(len(pasta_das_planilhas)) )
        window_veiculos_imoveis['-Progresso_texto-'].update( str( round(float(count) / int(len(pasta_das_planilhas)) * 100, 1) ) + '%' )
        window_veiculos_imoveis.refresh()
        
        if event == '-Encerrar-':
            return
        

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    # Especifique o caminho para o arquivo de trava
    lock_file_path = 'lancamentos_irpf.lock'
    
    # Verifique se outra instância está em execução
    if not create_lock_file(lock_file_path):
        p.alert(text="Outra instância já está em execução.")
        sys.exit(1)
        
    # Defina uma função para remover o arquivo de trava ao final da execução
    atexit.register(remove_lock_file, lock_file_path)
    
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    # sg.theme_previewer()
    # Layout da janela
    
    def janela_menu():  # layout da janela do menu principal
        layout_menu = [
            [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True)],
            [sg.Text('')],
            [sg.Text('Selecione o diretório com os arquivos .DBK originais:')],
            [sg.FolderBrowse('Pesquisar', key='-abrir_arq_1-'), sg.InputText(key='-pasta_arquivos_originais-', size=80, disabled=True)],
            [sg.Text('Selecione um diretório para salvar os arquivos .DBK editados e também a planilha de andamentos:')],
            [sg.FolderBrowse('Pesquisar', key='-abrir_arq_2-'), sg.InputText(key='-pasta_arquivos_editados-', size=80, disabled=True)],
            [sg.Text('')],
            [sg.Text('Selecione o lançamento que deseja inserir nos arquivos', justification='center')],
            [sg.Button('Ações e Fundo - Planilhas B3', key='-acoes_fundos-', size=30, border_width=1),
             sg.Button('Veículos e Imóveis', key='-veiculos_imoveis-', size=30, border_width=1),]
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Lançamentos IRPF', layout_menu, finalize=True)
    
    
    def janela_acoes_fundos():  # layout da janela do menu de ações e fundos
        layout_acoes_fundos = [
            [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True)],
            [sg.Text('')],
            [sg.Text('Selecione um diretório com as planilhas da B3')],
            [sg.FolderBrowse('Pesquisar', key='-Abrir2-'), sg.InputText(key='-pasta_final_b3-', size=80, disabled=True)],
            [sg.Text('Selecione a planilha com os dados dos clientes, a planilha deve conter as seguintes colunas apenas: "CPF" "NOME"')],
            [sg.FileBrowse('Pesquisar', key='-Abrir2-'), sg.InputText(key='-dados_b3-', size=80, disabled=True)],
            [sg.Text('')],
            [sg.Text('', key='-Mensagens-')],
            [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color='#f0f0f0')],
            [sg.Button('Iniciar', key='-Iniciar-', border_width=0),
             sg.Button('Encerrar', key='-Encerrar-', disabled=True, border_width=0),
             sg.Button('Abrir resultados', key='-Abrir resultados-', disabled=True, border_width=0),
             sg.Button('Abrir pasta das planilhas', key='-abrir_planilhas-', disabled=True, border_width=0)],
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Lançamentos Ações e Fundos', layout_acoes_fundos, finalize=True, modal=True)
    
    
    def janela_veiculos_imoveis():  # layout da janela do menu de veículos e imóveis
        layout_veiculos_imoveis = [
            [sg.Button('Ajuda', border_width=0), sg.Button('Log do sistema', border_width=0, disabled=True)],
            [sg.Text('')],
            [sg.Text('Selecione um diretório com as planilhas dos veículos e imóveis:')],
            [sg.FolderBrowse('Pesquisar', key='-Abrir2-'), sg.InputText(key='-planilhas_veiculos_imoveis-', size=80, disabled=True)],
            [sg.Text('')],
            [sg.Text('', key='-Mensagens-')],
            [sg.Text(size=6, text='', key='-Progresso_texto-'), sg.ProgressBar(max_value=0, orientation='h', size=(54, 5), key='-progressbar-', bar_color='#f0f0f0')],
            [sg.Button('Iniciar', key='-Iniciar-', border_width=0), sg.Button('Encerrar', key='-Encerrar-', disabled=True, border_width=0), sg.Button('Abrir resultados', key='-Abrir resultados-', disabled=True, border_width=0)],
        ]
        
        # guarda a janela na variável para manipula-la
        return sg.Window('Lançamentos Veículos e Imóveis', layout_veiculos_imoveis, finalize=True, modal=True)
    
    
    # inicia as variáveis das janelas
    window_menu, window_acoes_fundos, window_veiculos_imoveis = janela_menu(), None, None
    
    while True:
        # captura o evento e os valores armazenados na interface
        window, event, values = sg.read_all_windows()
        
        if event == sg.WIN_CLOSED:
            if window == window_veiculos_imoveis:  # if closing win 2, mark as closed
                window_veiculos_imoveis = None
            if window == window_acoes_fundos:  # if closing win 2, mark as closed
                window_acoes_fundos = None
            elif window == window_menu:  # if closing win 1, exit program
                break

        caminho_arquivo_original = values['-pasta_arquivos_originais-']
        caminho_arquivo_editados = values['-pasta_arquivos_editados-']

        if event == 'Ajuda':
            os.startfile('Manual do usuário - Lançamentos IRPF.pdf')
        
        elif event == 'Log do sistema':
            os.startfile('Log')
            
        elif event == '-acoes_fundos-':
            if caminho_arquivo_original:
                nao_tem = True
                for arquivo in os.listdir(caminho_arquivo_original):
                    if arquivo.endswith('.DBK'):
                        nao_tem = False
                if not nao_tem:
                    if caminho_arquivo_editados:
                        window_acoes_fundos = janela_acoes_fundos()
                        window_menu.Hide()
                        def run_script_thread():
                            window_acoes_fundos['-Iniciar-'].update(disabled=True)
                            window_acoes_fundos['-Encerrar-'].update(disabled=False)
                            window_acoes_fundos['-Abrir resultados-'].update(disabled=False)
                            # apaga qualquer mensagem na interface
                            window_acoes_fundos['-Mensagens-'].update('')
                            # atualiza a barra de progresso para ela ficar mais visível
                            window_acoes_fundos['-progressbar-'].update(bar_color=('#fca400', '#ffe0a6'))
                            
                            # Chama a função que executa o script
                            try:
                                run_acoes_fundos(window_acoes_fundos, event, dados_b3, final_folder, caminho_arquivo_original, caminho_arquivo_editados)
                            except Exception as erro:
                                time.sleep(1)
                                p.alert(erro)
                                
                            window_acoes_fundos['-Iniciar-'].update(disabled=False)
                            window_acoes_fundos['-Encerrar-'].update(disabled=True)
                            # apaga a porcentagem e a barra de progresso para a interface ficar mais limpa
                            window_acoes_fundos['-progressbar-'].update_bar(0)
                            window_acoes_fundos['-progressbar-'].update(bar_color='#f0f0f0')
                            window_acoes_fundos['-Progresso_texto-'].update('')
                            window_acoes_fundos['-Mensagens-'].update('')

                        while True:
                            # captura o evento e os valores armazenados na interface
                            event, values = window_acoes_fundos.read()
                            
                            try:
                                final_folder = values['-pasta_final_b3-']
                                dados_b3 = values['-dados_b3-']
                                download_folder = "C:\\Planilhas"
                            except:
                                final_folder = None
                                dados_b3 = None
                                download_folder = None
                                
                            if event == sg.WIN_CLOSED:
                                break
                            
                            elif event == 'Log do sistema':
                                os.startfile('Log')
                            
                            elif event == 'Ajuda':
                                os.startfile('Manual do usuário - Lançamentos IRPF.pdf')
                            
                            elif event == '-Iniciar-':
                                # Cria uma nova thread para executar o script
                                script_thread = Thread(target=run_script_thread)
                                script_thread.start()
                            
                            elif event == '-Abrir resultados-':
                                os.startfile(os.path.join(caminho_arquivo_editados))
                        
                        window_acoes_fundos.close()
                        window_menu.UnHide()
                    else:
                        p.alert(text='Selecione uma pasta para salvar os arquivos editados.')
                else:
                    p.alert(text='Não existe arquivo ".DBK" para editar na pasta selecionada.')
            else:
                p.alert(text='Selecione uma pasta com os arquivos para editar.')
            
        elif event == '-veiculos_imoveis-':
            if caminho_arquivo_original:
                nao_tem = True
                for arquivo in os.listdir(caminho_arquivo_original):
                    if arquivo.endswith('.DBK'):
                        nao_tem = False
                if not nao_tem:
                    if caminho_arquivo_editados:
                        window_veiculos_imoveis = janela_veiculos_imoveis()
                        window_menu.Hide()
                        def run_script_thread():
                            window_veiculos_imoveis['-Iniciar-'].update(disabled=True)
                            window_veiculos_imoveis['-Encerrar-'].update(disabled=False)
                            window_veiculos_imoveis['-Abrir resultados-'].update(disabled=False)
                            # apaga qualquer mensagem na interface
                            window_veiculos_imoveis['-Mensagens-'].update('')
                            # atualiza a barra de progresso para ela ficar mais visível
                            window_veiculos_imoveis['-progressbar-'].update(bar_color=('#fca400', '#ffe0a6'))
                            
                            # Chama a função que executa o script
                            try:
                                run_veiculos_moveis(window_veiculos_imoveis, event, pasta_planilhas_veiculos_imoveis, caminho_arquivo_original, caminho_arquivo_editados)
                            except Exception as erro:
                                time.sleep(1)
                                p.alert(erro)
                            
                            window_veiculos_imoveis['-Iniciar-'].update(disabled=False)
                            window_veiculos_imoveis['-Encerrar-'].update(disabled=True)
                            # apaga a porcentagem e a barra de progresso para a interface ficar mais limpa
                            window_veiculos_imoveis['-progressbar-'].update_bar(0)
                            window_veiculos_imoveis['-progressbar-'].update(bar_color='#f0f0f0')
                            window_veiculos_imoveis['-Progresso_texto-'].update('')
                            window_veiculos_imoveis['-Mensagens-'].update('')
                            
                        while True:
                            # captura o evento e os valores armazenados na interface
                            event, values = window_veiculos_imoveis.read()
                            
                            try:
                                pasta_planilhas_veiculos_imoveis = values['-planilhas_veiculos_imoveis-']
                            except:
                                pasta_planilhas_veiculos_imoveis = 'Desktop'
                                
                            if event == sg.WIN_CLOSED:
                                break
                            
                            elif event == 'Log do sistema':
                                os.startfile('Log')
                            
                            elif event == 'Ajuda':
                                os.startfile('Manual do usuário - Lançamentos IRPF.pdf')
                            
                            elif event == '-Iniciar-':
                                # Cria uma nova thread para executar o script
                                script_thread = Thread(target=run_script_thread)
                                script_thread.start()
                            
                            elif event == '-Abrir resultados-':
                                os.startfile(os.path.join(caminho_arquivo_editados))
                        
                        window_veiculos_imoveis.close()
                        window_menu.UnHide()
                    else:
                        p.alert(text='Selecione uma pasta para salvar os arquivos editados.')
                else:
                    p.alert(text='Não existe arquivo ".DBK" para editar na pasta selecionada.')
            else:
                p.alert(text='Selecione uma pasta com os arquivos para editar.')
    
    window.close()
    