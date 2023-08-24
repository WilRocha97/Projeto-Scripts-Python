# -*- coding: utf-8 -*-
import time
import fitz, re, os
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from datetime import datetime
from pathlib import Path
from pyautogui import alert


def escreve_relatorio_csv(texto, local, nome='Relatório', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local,  f"{nome} - complementar.csv"), 'a', encoding=encode)
    
    f.write(texto + '\n')
    f.close()


def ask_for_dir():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title='Selecione onde salvar a planilha',
    )
    
    return folder if folder else False


def ask_for_file():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title='Selecione o Relatório de Sindicatos',
        filetypes=[('PDF files', '*.pdf *')],
        initialdir=os.getcwd()
    )
    
    return file if file else False


def analiza():
    # pergunta qual PDF analizar
    relatorio = ask_for_file()
    # pergunta onde salvar a planilha com as informações
    final = ask_for_dir()
    
    # Abrir o pdf
    with fitz.open(relatorio) as pdf:
        
        # Para cada página do pdf
        for page in pdf:
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # procura quantas rubricas existem na página
                cadastros = re.compile(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-)\n(\d\d)\n(\d\d\d.\d\d\d.\d\d\d.\d\d\d)\n(.+)').findall(textinho)
                
                # para cada rubrica executa
                for cadastro in cadastros:
                    cnpj = cadastro[0] + cadastro[1]
                    ie = cadastro[2]
                    nome = cadastro[3]
                    print(f"{cnpj};{ie};{nome}")
                
                    escreve_relatorio_csv(f"{str(cnpj.replace(' ', ''))};{ie};{nome}", nome=relatorio, local=final)
        
            except:
                escreve_relatorio_csv(f"{page.number};Erro", nome=relatorio, local=final)
        return final
        

if __name__ == '__main__':
    inicio = datetime.now()
    analiza()
    alert(title='Gera relatório de sindicato', text='Relatóro gerado com sucesso.')