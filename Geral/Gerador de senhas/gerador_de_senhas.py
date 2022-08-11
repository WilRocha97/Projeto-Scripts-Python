from random import choice
import string
from pyautogui import prompt, alert
from tkinter import *


def configura():
    quantidade = prompt(title='Script incrível', text='Quantas senhas?')
    tamanho = prompt(title='Script incrível', text='Quantos caracteres?')
    
    # inicia o tkinter
    root = Tk()
    root.title('Script incrível')

    window_width = 300
    window_height = 125

    # find the center point
    center_x = int(root.winfo_screenwidth() / 2 - window_width / 2)
    center_y = int(root.winfo_screenheight() / 2 - window_height / 2)

    # set the position of the window to the center of the screen
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    # inicia as variáveis que serão armazenadas as opções do tkinter
    numero = IntVar()
    letra = IntVar()
    carac_espec = IntVar()
    
    # checkbox
    a = Checkbutton(root, text="Números", variable=numero)
    a.pack()
    # checkbox
    b = Checkbutton(root, text="Letras", variable=letra)
    b.pack()
    # checkbox
    c = Checkbutton(root, text="Caracteres especiais", variable=carac_espec)
    c.pack()
    
    # botão com comando para fechar o tkinter
    ok = Button(root, text="Ok", command=root.destroy)
    ok.pack()
    
    root.mainloop()
    
    return quantidade, tamanho, '!@#$%&*_?', '', 1, numero, letra, carac_espec


def gera_senhas():
    quantidade, tamanho, especiais, senha_anterior, indice, numero, letra, carac_espec = configura()
    
    while indice <= int(quantidade):
        
        valores = ''
        senha = ''
        
        if numero.get() == 1:
            valores += string.digits
        if letra.get() == 1:
            valores += string.ascii_letters
        if carac_espec.get() == 1:
            valores += especiais
            
        try:
            for b in range(int(tamanho)):
                caractere = choice(valores)
                senha += caractere
                valores = valores.replace(caractere, '')
            
            if senha != senha_anterior:
                senha_anterior = senha
                print(senha)
                indice += 1
                try:
                    f = open(str(f"Senhas {tamanho} dígitos.csv"), 'a', encoding='latin-1')
                except:
                    f = open(str(f"Senhas {tamanho} dígitos - auxiliar.csv"), 'a', encoding='latin-1')
                
                f.write(senha + '\n')
                f.close()
        except:
            return False
    
    return True


if __name__ == '__main__':
    while not gera_senhas():
        alert(title='Script incrível', text='Formato inválido!\n'
                                            'Apenas números: Máximo de 10 caracteres\n'
                                            'Apenas letras: Máximo de 26 caracteres\n'
                                            'Apenas caracteres especiais: Máximo de 9 caracteres\n'
                                            'Números com letras: Máximo de 36 caracteres\n'
                                            'Números com caracteres especiais: Máximo de 19 caracteres\n'
                                            'Letras com caracteres especiais: Máximo de 35 caracteres\n'
                                            'Tamanho máximo de 45 caracteres')
