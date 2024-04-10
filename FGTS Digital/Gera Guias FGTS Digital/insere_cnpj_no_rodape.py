import os, time
import shutil
import fitz
import re
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir

# pasta final
caminho_final = os.path.join('execução', 'Guias CNPJ completo')
os.makedirs(caminho_final, exist_ok=True)
documentos = ask_for_dir()


def add_text_to_pdf(guia, caminho, caminho_final, cnpj):
    x_text = 46  # Posição x do texto na página
    y_text = 830  # Posição y do texto na página
    
    # Abrir o PDF de entrada
    pdf_document = fitz.open(os.path.join(caminho, guia))
    
    # Selecionar a página desejada
    page = pdf_document[0]
    
    # Adicionar texto na página
    page.insert_text((x_text, y_text), cnpj, fontsize=12, overlay=True)
    
    # Salvar o PDF modificado
    pdf_document.save(os.path.join(caminho_final, guia))


# for cnpj in cnpjs:
for file in os.listdir(documentos):
    if not file.endswith('.pdf'):
        continue
    arq = os.path.join(documentos, file)
    doc = fitz.open(arq, filetype="pdf")
    
    texto_arquivo = ''
    for page in doc:
        texto = page.get_text('text', flags=1 + 2 + 8)
        texto_arquivo += texto
    
    # print(texto_arquivo)
    # time.sleep(55)
    
    cnpj = re.compile(r'(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)').search(file).group(1)
    
    add_text_to_pdf(file, documentos, caminho_final, cnpj)

