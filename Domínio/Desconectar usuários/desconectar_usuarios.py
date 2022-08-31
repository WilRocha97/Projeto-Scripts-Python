# -*- coding: utf-8 -*-
from os import listdir, path
from time import sleep
from tkinter.filedialog import askdirectory, Tk
import pyautogui as p
import cv2


def ask_for_dir(title='Abrir pasta'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False


def localiza_autorizados():
    for imagem in listdir(documentos):
        if p.locateOnScreen(path.join(documentos, imagem)):
            while p.locateOnScreen(path.join(documentos, imagem)):
                p.moveTo(p.locateCenterOnScreen(path.join(documentos, imagem)))
                local_mouse = p.position()
                p.click(int(local_mouse[0] - 63), local_mouse[1])
                sleep(0.5)


def run():
    while p.locateOnScreen(path.join('imgs', 'tela_de_carregamento.png'), confidence=0.9):
        sleep(1)
        
    p.click(p.locateCenterOnScreen(path.join('imgs', 'lista_usuarios.png'), confidence=0.9), button='left', clicks=2)
    p.click(p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9))
    
    p.hotkey('alt', 't')
    sleep(2)
    
    while not p.locateOnScreen(path.join('imgs', 'seta_baixo_limite.png')):
        localiza_autorizados()
        
        p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_baixo.png'), confidence=0.9), button='left', clicks=10)
        
        localiza_autorizados()
        
    p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_cima.png'), confidence=0.9), button='left', clicks=50)
    
    # if p.locateCenterOnScreen(os.path.join('imgs', 'desconectar.png'), confidence=0.9):
    #     p.hotkey('alt', 'd')
    
    
if __name__ == '__main__':
    documentos = ask_for_dir()
    while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.5):
        sleep(1)
        
    while p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.5):
        run()
        sleep(30)
    
    p.alert(text='Robô finalizado.')
    