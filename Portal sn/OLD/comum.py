# -*- coding: utf-8 -*-
from time import sleep
import pyautogui as a
import os, sys

def find_img(img, path='.', conf=1):
    path = os.path.join(path, 'imgs', img)
    return a.locateOnScreen(path, confidence=conf)


def click_img(img, path='.', conf=1, tempo=0):
    try:
        img = os.path.join(path, 'imgs', img)
        x, y = a.locateCenterOnScreen(img , confidence=conf)
        a.click(x, y)
        sleep(tempo)
    except:
        return False

    return True


def wait_img(img, path='.', conf=1, timeout=10):
    for i in range(timeout):
        if find_img(img=img, path=path, conf=conf): return True
        sleep(1)
    else:
        return False


def escrever_relatorio_csv(texto, end='\n'):
    try:
        file = open('resumo.csv', 'a')
    except:
        file = open('resumo_erro.csv', 'a')
    file.write(texto + end)
    file.close()


def login_sn(cnpj):
    print('>>> logando com', cnpj)

    url_login = 'http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao' 
    os.system(f'start chrome /new-window {url_login}')

    if not wait_img('input_cnpj.png', path='..', conf=0.9):
        return False, 'Timeout - Não carregou pagina login'

    a.write(cnpj)
    sleep(0.2)
    a.press('enter')

    if not wait_img('btn_emitir.png', path='..', conf=0.9):
        return False, 'Timeout - Não carregou pagina inicial'

    sleep(0.2)
    return True, ''