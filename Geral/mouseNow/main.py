# -*- coding: utf-8 -*-
import pyautogui as a
from threading import Thread
from tkinter import Tk, filedialog, messagebox
from socket import socket, AF_INET, SOCK_STREAM
import eel
from os.path import exists, join


_ATIVO = True
_CAMINHO = ''

def porta_em_uso(porta):
    with socket(AF_INET, SOCK_STREAM) as s:
        return s.connect_ex(('localhost', porta)) == 0


eel.init('web', allowed_extensions=['.css', '.html', '.js'])


@eel.expose
def handle_exit(ar1, ar2):
    import sys
    global _ATIVO
    _ATIVO = False
    print('Finalizado')
    sys.exit(0)


def inicia_rotina():
    global _ATIVO
    while _ATIVO:
        try:
            x, y = a.position()
            color = a.screenshot().getpixel((x, y))
            RGB = ','.join([str(c).rjust(3) for c in color])
            eel.atualiza_coordenadas(x, y, RGB)
        except IndexError:
            pass


@eel.expose
def salvar_imagem(x, y):
    global _CAMINHO

    caminho = _CAMINHO if _CAMINHO else '.'
    x = [int(i) for i in x.split(' , ')]
    y = [int(i) for i in y.split(' , ')]
    w = y[0] - x[0]
    h = y[1] - x[1]
    if w < 0 and h < 0:
        i, j = y
        w, h = abs(w), abs(h)
    elif w < 0 and h >= 0:
        i, j = y[0], x[1]
        w = abs(w)
    elif w >= 0 and h < 0:
        i, j = x[0], y[1]
        h = abs(h)
    else:
        i, j = x

    total = 0
    while True:
        nome = f'imagem ({total}).png' if total else 'imagem.png'
        if not exists(join(caminho, nome)):
            try:
                a.screenshot(join(caminho, nome), region=(i, j, w, h))
                break
            except Exception as e:
                print(e)
                return False
        else:
            total += 1

    return nome


@eel.expose
def busca_caminho():
    global _CAMINHO

    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    _CAMINHO = filedialog.askdirectory()
    lista = [x for x in _CAMINHO.split('/') if x]
    texto = '/'.join(lista)
    if len(lista) > 4:
        texto = '/'.join([lista[0], lista[1], '...', lista[-2], lista[-1]])
    return texto


@eel.expose
def get_position():
    job = Thread(target=inicia_rotina)
    job.start()


porta = 62110
maximo = 62115
while porta <= maximo:
    try:
        if not porta_em_uso(porta):
            #eel.spawn(inicia_rotina)
            eel.start('index.html', port=porta, size=(440,362), close_callback=handle_exit)
            break
    except Exception as e:
        print(e)
    porta += 1
else:
    root = Tk()
    root.withdraw()
    messagebox.showwarning("Erro", "Falha ao abrir o programa.")