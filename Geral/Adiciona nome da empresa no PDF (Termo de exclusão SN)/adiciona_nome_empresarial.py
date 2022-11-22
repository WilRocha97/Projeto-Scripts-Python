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
        '''print(texto)
        time.sleep(55)'''
        try:
            regex_termo = re.compile(r'(.+)\nNome Empresarial: \n')
            resultado = regex_termo.search(texto).group(1)
        except:
            try:
                regex_termo = re.compile(r'Nome Empresarial: (.+)')
                resultado = regex_termo.search(texto).group(1)
            except:
                regex_termo = re.compile(r' (.+)\nNome Empresarial:\n')
                resultado = regex_termo.search(texto).group(1)
        
        doc.close()
        new = os.path.join(documentos, f'{file.replace(".pdf", " - " + resultado)}.pdf')
        shutil.move(arq, new)
        break
