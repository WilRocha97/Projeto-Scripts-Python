# -*- coding: utf-8 -*-
from unicodedata import combining, normalize
from msoffcrypto import OfficeFile
from openpyxl import load_workbook
from datetime import datetime
from os import remove, walk
import pdfplumber


def normaliza_string(string):
    str_nova = ''.join(l for l in normalize('NFKD', string) 
        if not combining(l)).replace("'", "").replace("\n", " ").upper().strip()
    
    return str_nova


def analisa_planilha(planilha):
    file = OfficeFile(open(planilha, "rb"))
    file.load_key(password="Veiga237")
    file.decrypt(open("temp.xlsx", "wb"))
    arquivo = load_workbook("temp.xlsx")
    nome_tabela = arquivo.get_sheet_names()[0]
    tabela = arquivo.get_sheet_by_name(nome_tabela)
    lista = []
    for linha in tabela.rows:
        nome = str(linha[0].value).strip() if linha[0].value else ''
        cpf = str(linha[1].value).strip() if linha[1].value else ''
        acesso = str(linha[2].value).strip() if linha[2].value else ''
        senha = str(linha[3].value).strip() if linha[3].value else ''
        dados = [nome, cpf, acesso, senha]
        if any(dados):
            lista.append(dados)
    remove('temp.xlsx')

    return lista



# VERIFICAR OS ARQUIVOS .PDF ANTES DE EXECUTAR O PROGRAMA
# CASO SEJA UM ARQUIVO ESCANEADO, SOLICITAR UM NOVO ARQUIVO 
if __name__ == '__main__':
    inicio = datetime.now()
    lista = analisa_planilha('IRPF 2018.xlsx') # COLOCAR O NOME DA PLANILHA AQUI
    arquivo = open('Dados.py', 'w')
    arquivo.write('empresas = [\n')
    for raiz, sub, arquivos in walk('.'):
        for arq in arquivos:
            nome = arq[:-4]+'.txt'
            extensao = arq[-4:]
            if extensao != '.pdf':
                continue
            nota = open(nome, 'w', encoding='utf-8')
            with pdfplumber.open(arq) as pdf:
                for pagina in pdf.pages[:]:
                    texto = pagina.extract_text(y_tolerance=14) # AJUSTAR TOLERANCIA CASO NECESSARIO
                    if not texto:
                        continue
                    nota.write(texto)
                    nota.write('\n')

                    dados = [x.split(' ') for x in texto.split('\n')]
                    for d in dados:
                        if len(d) <= 3:
                            print('>>> Dados invalidos: ', d)
                            continue
                        cpf = d[0].replace('.', '').replace('-', '')
                        recibo = d[3].replace('.', '')
                        if not recibo.isnumeric():
                            continue
                        for registro in lista:
                            if registro[1] == cpf:
                                registro[0] = normaliza_string(registro[0])
                                linha = '('
                                for dado in registro:
                                    linha += "'"+dado+"'"+","
                                linha += "),"
                                arquivo.write(linha+'\n')
            nota.close()
            remove(nome) # Remove os arquivos convertidos 
    arquivo.write(']')
    arquivo.close()

    print('>>> Programa finalizado em', (datetime.now() - inicio))
