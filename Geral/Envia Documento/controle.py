# -*- coding: utf-8 -*-
import os, pyautogui as p
from pathlib import Path

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img

try:
    os.makedirs(os.path.join('ignore', 'controle'), exist_ok=True)
    while 1 > 0:
        if _find_img('numero_invalido.png', conf=0.9):
            with open(os.path.join('ignore', 'controle', 'Número inválido.txt'), 'w') as arquivo:
                arquivo.write('Número inválido')
except:
    p.alert(text='Controle encerrado')
