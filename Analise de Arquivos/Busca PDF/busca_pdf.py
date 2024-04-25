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
pasta = os.path.join('execução', 'Documentos')
os.makedirs(pasta, exist_ok=True)

# Percorre a pasta e suas subpastas
contador = 1
comecar = False
for pasta_atual, subpastas, arquivos in os.walk(pasta_principal):
    # Verifica se a subpasta atual é a que você deseja ignorar
    if subpasta_a_ignorar in subpastas:
        subpastas.remove(subpasta_a_ignorar)  # Remove a subpasta da lista para ignorá-la
    
    """if not comecar:
        if not pasta_atual.startswith('T:/RO'):
            continue
        else:
            comecar = True"""
        
    # Agora você pode processar os arquivos na pasta atual normalmente
    for file in arquivos:
        
        caminho_completo = os.path.join(pasta_atual, file)
        
        if not file.endswith('.pdf'):
            continue
        arq = os.path.join(pasta_atual, file)
        try:
            doc = fitz.open(arq, filetype="pdf")
            encontrou = False
        except:
            arq = arq.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')
            pasta_atual = pasta_atual.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')
            try:
                _escreve_relatorio_csv(f'{arq};Documento corrompido', nome='Buscar_doc')
                continue
            except:
                _escreve_relatorio_csv(f'{pasta_atual};Possuí documento corrompido', nome='Buscar_doc')
                continue
                
        print(caminho_completo)
        try:
            arq = arq.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')
            pasta_atual = pasta_atual.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')
            for page in doc:
                texto = page.get_text('text', flags=1 + 2 + 8)
                
                regexes = [r'Comprovante de Inscrição e de Situação Cadastral\nContribuinte,  \nConfira os dados de Identificação da Pessoa Jurídica e, se houver qualquer divergência, providencie junto à \nRFB a sua atualização cadastral',
                           r'Comprovante de Inscrição e de Situação Cadastral\nContribuinte,\nConfira os dados de Identificação da Pessoa Jurídica e, se houver qualquer divergência, providencie junto à\nRFB a sua atualização cadastral',
                           r'Comprovante de Inscrição e de Situação Cadastral\nContribuinte,\nConfira os dados de Identificação da Pessoa Jurídica e, se houver qualquer divergência, providencie junto à\nRFB a sua atualização cadastral',
                           r'REPÚBLICA FEDERATIVA DO BRASIL\s+CADASTRO NACIONAL DA PESSOA JURÍDICA'
                ]
                for regex in regexes:
                    resultado = re.compile(regex).search(texto)
                    if not resultado:
                        continue
                    else:
                        # print(texto)
                        # time.sleep(55)
                        doc.close()
                        new = os.path.join(pasta, file.replace('.pdf', f'{contador}.pdf'))
                        shutil.move(arq, new)
                        _escreve_relatorio_csv(f'{arq};{file.replace(".pdf", f"{contador}.pdf")}', nome='Buscar_doc')
                        contador += 1
                        encontrou = True
                        break
                
                # sem for consultar documentos com mais de uma página usa esse
                """if encontrou:
                    break"""
                # se for consultar em documentos com apenas uma página usa esse
                break
        except:
            arq = arq.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')
            pasta_atual = pasta_atual.replace("–", "-").replace("$", "S").replace(u'\u2022', ' ').replace(u'\u201c', ' ')

            try:
                _escreve_relatorio_csv(f'{arq};Documento com senha', nome='Buscar_doc')
            except:
                _escreve_relatorio_csv(f'{pasta_atual};Possuí documento com senha', nome='Buscar_doc')
        