import os, time
import shutil
import fitz
import re
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir

documentos = ask_for_dir()

# for cnpj in cnpjs:
for file in os.listdir(documentos):
    if not file.endswith('.pdf'):
        continue
    arq = os.path.join(documentos, file)
    doc = fitz.open(arq, filetype="pdf")
    
    for page in doc:
        texto = page.get_text('text', flags=1 + 2 + 8)

        regex_termo = re.compile(r'PROCURADORIA-GERAL DA FAZENDA NACIONAL\n(\d\d/\d\d/\d\d\d\d)')
        resultado = regex_termo.search(texto)
        
        data = resultado.group(1)
        data_separada = data.split('/')
        
        dia = data_separada[0]
        mes = data_separada[1]
        ano = data_separada[2]
        
        pasta_nova = os.path.join('Execução', ano, f'{mes}-{ano}', f'{dia}-{mes}-{ano}')
        os.makedirs(pasta_nova, exist_ok=True)
        doc.close()
        shutil.move(arq, os.path.join(pasta_nova, file))
        break
