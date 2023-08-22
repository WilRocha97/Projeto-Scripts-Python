# -*- coding: utf-8 -*-
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from tkinter import *
from tkinter import END
from time import sleep
from os import path, getcwd
import fitz
import re


# Para ocultar o prompt acrescenta --noconsole ou um único arquivo zipado --onefile


def salva_csv(texto, caminho='.'):
    nome = path.join(caminho, 'Analisa Sindicato.csv')
    arq = open(nome, 'a')
    arq.write(texto + '\n')
    arq.close()
    sleep(0.2)


def busca_dados(title='Abrir arquivo', initialdir='ignore'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    filetypes = (
        ('Documento do Adobe Acrobat', '*.pdf'),
    )
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    root.destroy()
    root.mainloop()
    
    if file:
        return file
    else:
        return False


def analisa_relatorio(file):
    if file:
        file_name = re.compile(r'.+/(.+\.pdf)').search(file)
        file_name = file_name.group(1)
        # inicia o tkinter
        root = Tk()
        root.title('Analisa sindicato')
        # traz a janela para frente
        root.attributes("-topmost", True)
        
        window_width = 300
        window_height = 200
        
        # find the center point
        center_x = int(root.winfo_screenwidth() / 2 - window_width / 2)
        center_y = int(root.winfo_screenheight() / 2 - window_height / 2)
        
        # set the position of the window to the center of the screen
        root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # botão com comando para fechar o tkinter
        final = Label(root, text=f'Arquivo selecionado: {file_name}\nSelecione de qual sistema o arquio foi gerado.')
        final.pack()
        
        sistema1 = intVar()
        sistema2 = intVar()
        sistemas = Checkbutton(root, text='Domínio', variable=sistema1)
        sistemas.pack()
        
        sistemas2 = Checkbutton(root, text='DP Cuca', variable=sistema2)
        sistemas2.pack()
        
        final2 = Label(root, text='Pressione "Ok" para continuar\ne aguarde a mensagem de finalização.')
        final2.pack()
        
        ok = Button(root, text="Ok", command=root.destroy)
        ok.pack()
        
        root.mainloop()

        if sistema1.get() == 1:
            paginas = []
            caminho = '/'.join(file.split('/')[:-1])
            salva_csv('Data;Código;Empresa;CNPJ;Sindicato;Total Calculado;Total Informado', caminho)
            
            regex_empresa = re.compile(r'Local de trabalho\n(.+\d) - (.+)\n(.+)')
            regex_cnpj_emp = re.compile(r'(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)\nRubrica')
            regex_sindicato = re.compile(r'Sindicato:\n.+ - (.+)')
            regex_valor = re.compile(r'Total da empresa:\n(.+)\n(.+)')
            doc = fitz.open(file)
            
            for index, page in enumerate(doc, 1):
                texto = page.get_text('text', flags=1 + 2 + 8)
                empresa = regex_empresa.search(texto)
                cnpj_emp = regex_cnpj_emp.search(texto)
                sindicato = regex_sindicato.search(texto)
                valor = regex_valor.search(texto)
                
                if all([empresa, cnpj_emp, valor]):
                    cod, razao, data = empresa.groups()
                    cnpj_emp = cnpj_emp.group(1)
                    sindicato = sindicato.group(1)
                    calculado, informado = valor.groups()
                    
                    registro = ';'.join([data, cod, razao, cnpj_emp, sindicato, calculado, informado])
                    salva_csv(registro, caminho)
                else:
                    paginas.append(str(index))
                    
            doc.close()
            return ', '.join(paginas)
        
        if sistema2.get() == 1:
            paginas = []
            caminho = '/'.join(file.split('/')[:-1])
            salva_csv('Data;Código;Empresa;CNPJ;Sindicato;CNPJ;Funcionarios;Remuneração;Contribuição;Controle', caminho)
        
            regex_empresa = re.compile(r'(.+)\nEmpresa:(\d{4})\s+(.*)\n(.*)\n')
            regex_cnpj_emp = re.compile(r'CNPJ\/CEI:(.+)')
            regex_cnpj_sind = re.compile(r'CNPJ\/CPF:(.+)')
            regex_funcionario = re.compile(r'Funcionários Contribuintes\n\s*(.+)\n\s*(.+)\n\s*(.+)\n')
            doc = fitz.open(file)
        
            for index, page in enumerate(doc, 1):
                texto = page.getText('text', flags=1 + 2 + 8)
                empresa = regex_empresa.search(texto)
                cnpj_emp = regex_cnpj_emp.search(texto)
                cnpj_sind = regex_cnpj_sind.search(texto)
                funcionario = regex_funcionario.search(texto)
            
                if all([empresa, cnpj_emp, cnpj_sind, funcionario]):
                    data, cod, razao, sindicato = empresa.groups()
                    cnpj_emp = cnpj_emp.group(1)
                    cnpj_sind = cnpj_sind.group(1)
                    total, remuneracao, contribuicao = funcionario.groups()
                
                    registro = ';'.join([data, cod, razao, cnpj_emp, sindicato, cnpj_sind, total, remuneracao, contribuicao])
                    salva_csv(registro, caminho)
                else:
                    paginas.append(str(index))

            doc.close()
            return ', '.join(paginas)
    
    else:
        return False


def iniciar():
    file = busca_dados()
    try:
        paginas = analisa_relatorio(file)

        file_result = re.compile(r'(.+/).+\.pdf').search(file)
        file_result = file_result.group(1)
        
        # inicia o tkinter
        root = Tk()
        root.title('Analisa sindicato')
        # traz a janela para frente
        root.attributes("-topmost", True)
        
        window_width = 400
        window_height = 150
        
        # find the center point
        center_x = int(root.winfo_screenwidth() / 2 - window_width / 2)
        center_y = int(root.winfo_screenheight() / 2 - window_height / 2)
        
        # set the position of the window to the center of the screen
        root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # botão com comando para fechar o tkinter
        final = Label(root, text='Planilha "Analisa Sindicato" gerada com sucesso no seguinte diretório:')
        final.pack()
        
        texto = Text(root, height=5)
        texto.pack()
        texto.insert(END, file_result)

        ok = Button(root, text="Ok", command=root.destroy)
        ok.pack()
        
        root.mainloop()
        
        return paginas or True
    
    except Exception as e:
        print(e)
        
        # inicia o tkinter
        root = Tk()
        root.title('Analisa sindicato')
        # traz a janela para frente
        root.attributes("-topmost", True)
        
        window_width = 300
        window_height = 100
        
        # find the center point
        center_x = int(root.winfo_screenwidth() / 2 - window_width / 2)
        center_y = int(root.winfo_screenheight() / 2 - window_height / 2)
        
        # set the position of the window to the center of the screen
        root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # botão com comando para fechar o tkinter
        final = Label(root, text='Aplicação finalizada')
        final.pack()
        
        ok = Button(root, text="Ok", command=root.destroy)
        ok.pack()
        
        root.mainloop()
        
        return False


if __name__ == '__main__':
    iniciar()
