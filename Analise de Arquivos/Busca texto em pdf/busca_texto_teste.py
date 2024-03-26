import os, time
import shutil
import fitz
import re
from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv

# Pasta principal que você deseja percorrer
pasta_principal = ask_for_dir()

# Subpasta que você deseja ignorar
subpasta_a_ignorar = 'RELATÓRIO DE PESQUISAS'

# pasta final
pasta = os.path.join('documentos', 'resultados')
os.makedirs(pasta, exist_ok=True)

# Percorre a pasta e suas subpastas
contador = 1
for pasta_atual, subpastas, arquivos in os.walk(pasta_principal):
    # Verifica se a subpasta atual é a que você deseja ignorar
    if subpasta_a_ignorar in subpastas:
        subpastas.remove(subpasta_a_ignorar)  # Remove a subpasta da lista para ignorá-la
    
    # Agora você pode processar os arquivos na pasta atual normalmente
    for file in arquivos:
        caminho_completo = os.path.join(pasta_atual, file)
        
        if not file.endswith('.pdf'):
            continue
        arq = os.path.join(pasta_atual, file)
        try:
            doc = fitz.open(arq, filetype="pdf")
        except:
            arq = arq.replace("–", "-").replace("$", "S")
            pasta_atual = pasta_atual.replace("–", "-").replace("$", "S")
            try:
                _escreve_relatorio_csv(f'{arq};Documento corrompido')
                continue
            except:
                _escreve_relatorio_csv(f'{pasta_atual};Possuí documento corrompido')
                continue
                
        print(caminho_completo)
        try:
            for page in doc:
                texto = page.get_text('text', flags=1 + 2 + 8)
            
                regex_termo = re.compile(r'Ficha Cadastral Simplificada. Documento certificado por JUNTA COMERCIAL DO ESTADO DE SÃO PAULO.')
        
                resultado = regex_termo.search(texto)
                if not resultado:
                    continue
                else:
                    # print(texto)
                    # time.sleep(55)
                    doc.close()
                    new = os.path.join(pasta, file.replace('.pdf', f'{contador}.pdf'))
                    shutil.move(arq, new)
                    _escreve_relatorio_csv(f'{arq};{file.replace(".pdf", f"{contador}.pdf")}')
                    contador += 1
                    break
        except:
            _escreve_relatorio_csv(f'{arq};Documento com senha')