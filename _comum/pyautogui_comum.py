# -*- coding: utf-8 -*-
from datetime import datetime
from time import sleep
import pyautogui as a
import os


# Foca na janela que tenha 'title' no meio do título da janela
# e maximiza a janela caso 'maximize' seja True
# Retorna True e '' caso ache somente uma janela
# Retorna False e uma mensagem caso ache nenhuma ou mais de uma janela
def focus(title, maximize=True, ignore_erro=False):
    win = a.getWindowsWithTitle(title)
    
    if len(win) == 0:
        return False, 'Nenhuma janela encontrada'

    if len(win) > 1:
        return ignore_erro, 'Mais de uma janela encontrada'

    if maximize:
        win[0].maximize()
    win[0].activate()
    sleep(0.2)
    return True, 'Janela em foco'
_focus = focus


# Procura pela imagem img que atenda ao nível de correspondência conf
# Retorna uma tupla com os valores (x, y, altura, largura) caso ache a img
# Retorna None caso não ache a img
def find_img(img, pasta='imgs', conf=1.0):
    try:
        path = os.path.join(pasta, img)
        return a.locateOnScreen(path, confidence=conf)
    except:
        return False
_find_img = find_img


# Espera pela imagem 'img' que atenda ao nível de correspondência 'conf'
# Até que o 'timeout' seja excedido ou indefinidamente para 'timeout' negativo
# Retorna uma tupla com os valores (x, y, altura, largura) caso ache a img
# Retorna None caso não ache a imagem ou exceda o 'timeout'
def wait_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, debug=False):
    if debug:
        print('\tEsperando', img)

    aux = 0
    while True:
        box = find_img(img, pasta, conf=conf)
        if box:
            return box
        sleep(delay)

        if timeout < 0:
            continue
        if timeout == aux:
            break
        aux += 1

    return None
_wait_img = wait_img


# Procura e clica 'clicks' vezes com o botão 'button' na imagem 'img'
# que atenda ao nível de correspondência conf
# Retorna True caso ache a img
# Retorna False caso não ache a imagem ou exceda o 'timeout'
def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    try:
        for i in range(timeout):
            box = a.locateCenterOnScreen(img, confidence=conf)
            if box:
                a.click(a.locateCenterOnScreen(img, confidence=conf), button=button, clicks=clicks)
                return True
            sleep(delay)
        else:
            return False
    except:
        return False
_click_img = click_img


def click_position_img(img, operacao, pixels_x=0, pixels_y=0, pasta='imgs', conf=1.0, clicks=1):
    img = os.path.join(pasta, img)
    a.moveTo(a.locateCenterOnScreen(img, confidence=conf))
    local_mouse = a.position()
    if operacao == '+':
        a.click(int(local_mouse[0] + int(pixels_x)), int(local_mouse[1] + int(pixels_y)), clicks=clicks)
        return True
    if operacao == '-':
        a.click(int(local_mouse[0] - int(pixels_x)), int(local_mouse[1] - int(pixels_y)), clicks=clicks)
        return True
_click_position_img = click_position_img


# Exibe caixa de texto pedindo pela competência informando o modelo 'printable'
# e checa se a entrada satisfaz o modelo definido em 'strptime'
# Retorna a competencia em string em caso de sucesso
# Retorna false caso a caixa de texto seja fechada


def get_comp(printable, strptime, subject='competencia'):
    text = base = f'Digite {subject} no modelo {printable}:'

    while True:
        comp = a.prompt(title=f'Qual {subject}', text=text)
        if comp is None:
            return False

        try:
            datetime.strptime(comp, strptime)
            break
        except ValueError:
            text = f'{subject} não atende ao padrão\n{base}'
    return comp
_get_comp = get_comp
