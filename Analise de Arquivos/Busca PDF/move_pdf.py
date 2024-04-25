import os, time
import shutil
import fitz
import re
from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _indice, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def run():
    pasta_dos_arquivos = ask_for_dir()
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)
        indice, caminho, arquivo = empresa
        print(caminho)
        
        lista_caminho = caminho.split('\\')
        indice_ultimo_item = len(lista_caminho) - 1
        caminho = caminho.replace(f'\\{lista_caminho[indice_ultimo_item]}', '')
        
        print(caminho)
        print(arquivo)
        
        try:
            shutil.move(os.path.join(pasta_dos_arquivos, arquivo), os.path.join(caminho, arquivo))
            _escreve_relatorio_csv(f'{indice};{caminho};{arquivo}', nome='Mover_doc')
        except:
            _escreve_relatorio_csv(f'{indice};{caminho};{arquivo};Erro ao mover o arquivo', nome='Mover_doc')
        


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    
    if index is not None:
        run()
        