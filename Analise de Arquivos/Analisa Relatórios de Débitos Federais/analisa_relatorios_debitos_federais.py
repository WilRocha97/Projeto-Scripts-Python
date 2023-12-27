import os, time
import shutil
import fitz
import re
from decimal import Decimal
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _remove_espacos


def coleta_dados(doc):
    debitos = []
    devedor = False
    id_cliente = ''
    for page in doc:
        texto = page.get_text('text', flags=1 + 2 + 8)
        # print(texto)
        
        print(f'Analisando página {page.number + 1}')
        try:
            id_cliente = re.compile(r'CNPJ: (\d\d\.\d\d\d\.\d\d\d\/\d\d\d\d-\d\d)').search(texto).group(1)
        except:
            try:
                id_cliente = re.compile(r'CPF: (\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(texto).group(1)
            except:
                pass
            
        debitos_por_pagina = re.compile(r'(.+)\n(.+)\n(.+)\n(.+,.+)\n(.+,.+)\nDEVEDOR\n').findall(texto)
        
        if not debitos_por_pagina:
            continue
        
        else:
            devedor = True
            for debito_da_lista in debitos_por_pagina:
                debitos.append(debito_da_lista)
    
    if not devedor:
        return False
    
    analisa_dados(id_cliente, debitos)
    return True

    
def analisa_dados(id_cliente, debitos):
    debitos_anterior = ''
    contador = 0
    datas = ''
    valor_total = 0
    
    debitos_ordenados = sorted(debitos)

    for count, debito in enumerate(debitos_ordenados):
        print(f'Analisando débito em {debito[0]}')
        if count == 0:
            contador += 1
            datas += debito[1] + ', '
            valor_total = valor_total + Decimal(float(str(debito[4]).replace('.', '').replace(',', '.')))
            debitos_anterior = debito[0]
            continue
            
        elif debito[0] == debitos_anterior:
            contador += 1
            datas += debito[1] + ', '
            valor_total = valor_total + Decimal(float(str(debito[4]).replace('.', '').replace(',', '.')))
            debitos_anterior = debito[0]
            continue
        
        debitos_anterior = _remove_espacos(debitos_anterior)
        valor_total = round(valor_total, 2)
        valor_total = str(valor_total).replace('.', ',')
        _escreve_relatorio_csv(f'{id_cliente};{debitos_anterior};{contador} Débitos;{valor_total};{datas}', nome='Relatório Devedores')
        time.sleep(2)
        contador = 1
        datas = debito[1] + ', '
        valor_total = Decimal(float(str(debito[4]).replace('.', '').replace(',', '.')))
        debitos_anterior = debito[0]
    
    if contador != 0:
        valor_total = round(valor_total, 2)
        valor_total = str(valor_total).replace('.', ',')
        _escreve_relatorio_csv(f'{id_cliente};{debitos_anterior};{contador} Débitos;{valor_total};{datas}', nome='Relatório Devedores')


@_time_execution
def run():
    # pasta final
    pasta_recorrentes = os.path.join('Execução', 'Devedor recorrente')
    pasta = os.path.join('Execução', 'Relatórios')
    os.makedirs(pasta_recorrentes, exist_ok=True)
    os.makedirs(pasta, exist_ok=True)
    documentos = ask_for_dir()
    
    # for cnpj in cnpjs:
    for file in os.listdir(documentos):
        if not file.endswith('.pdf'):
            continue
        arq = os.path.join(documentos, file)
        doc = fitz.open(arq, filetype="pdf")
        print('\n' + file)
        
        coletou_dados = coleta_dados(doc)
        
        if coletou_dados:
            doc.close()
            new = os.path.join(pasta_recorrentes, file)
            shutil.move(arq, new)
        
    _escreve_header_csv('CNPJ;RECEITA;QUANTIDADE;TOTAL DO DÉBITO;COMPETÊNCIAS', nome='Relatório Devedores')
    
    
if __name__ == '__main__':
    run()
    