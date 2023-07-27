# -*- coding: utf-8 -*-
import time
from time import sleep
from pyperclip import copy
from pyautogui import press, click, hotkey, alert, prompt, mouseInfo

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _wait_img, _click_img
from comum_comum import _time_execution


def inserir():
    while _find_img('em_branco.png', conf=0.9):
        press('enter')
        _wait_img('K915.png', conf=0.9, timeout=-1)
        while not _find_img('campo_obrigatorio.png', conf=0.9):
            press('tab')
            time.sleep(0.5)
        _click_img('campo_obrigatorio.png', conf=0.9)
        time.sleep(0.2)
        copy('BALANÇO MENSAL ESTIMATIVA / SUSPENSÃO')
        copy('BALANÇO MENSAL ESTIMATIVA / SUSPENSÃO')
        hotkey('ctrl', 'v')
        time.sleep(0.5)
        press('enter', presses=2, interval=0.5)
        time.sleep(1)
        press('down')
        time.sleep(1)
        click(1285, 598, clicks=20)
            

@_time_execution
def run():
    _wait_img('em_branco.png', conf=0.9, timeout=-1)
    inserir()
    print()


if __name__ == '__main__':
    run()
