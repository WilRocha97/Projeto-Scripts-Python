# -*- coding: utf-8 -*-
import time
from time import sleep
from pyperclip import copy
from pyautogui import press, hotkey, write, click, alert, prompt, getWindowsWithTitle

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _wait_img
from comum_comum import _time_execution


def inserir():
    while _find_img('EmBranco.png', conf=0.9):
        press('enter')
        _wait_img('K915.png', conf=0.9, timeout=-1)
        press('tab', presses=5, interval=0.5)
        time.sleep(0.2)
        copy('BALANÇO MENSAL ESTIMATIVA / SUSPENSÃO')
        hotkey('ctrl', 'v')
        time.sleep(0.2)
        press('enter', presses=2, interval=0.5)
        time.sleep(1)
        press('down')
        time.sleep(1)


@_time_execution
def run():
    _wait_img('EmBranco.png', conf=0.9, timeout=-1)
    inserir()
    print()


if __name__ == '__main__':
    run()
