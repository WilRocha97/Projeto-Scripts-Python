# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv


def envia_email():
    p.getWindowsWithTitle('Cucmail')[0].maximize()
    
    while _find_img('PDF.png', 0.99):
        if _find_img('EmailNaoPodeSerEnviado.png', 0.95):
            _click_img('EmailNaoPodeSerEnviado.png', 0.9)
            p.press('enter')
            _wait_img('ChromeAberto.png', 0.9)
            time.sleep(3)
            p.hotkey('alt', 'f4')
            time.sleep(1)
            p.click(875, 142)
            time.sleep(1)
            _click_img('ExcluirEmail.png', 0.9)
            _wait_img('ConfirmarExclusao.png', 0.9)
            time.sleep(1)
            p.press('s')
            
        while not _find_img('EmailEnviado.png', 0.9):
            p.click(875, 142)
            time.sleep(1)

            if _find_img('EmailNaoPodeSerEnviado.png', 0.95):
                _click_img('EmailNaoPodeSerEnviado.png', 0.9)
                p.press('enter')
                _wait_img('ChromeAberto.png', 0.9)
                time.sleep(3)
                p.hotkey('alt', 'f4')
                time.sleep(1)
                p.click(875, 142)
                time.sleep(1)
                _click_img('ExcluirEmail.png', 0.9)
                _wait_img('ConfirmarExclusao.png', 0.9)
                time.sleep(1)
                p.press('s')
                break

            if _find_img('EmailEnviado.png', 0.9):
                _click_img('EmailEnviado.png', 0.9)
                p.press('enter')

            _click_img('Enviar.png', 0.9, clicks=3)
            time.sleep(1)
            while not _find_img('EmailEnviado.png', 0.9):
                time.sleep(2)

                if _find_img('EmailNaoPodeSerEnviado.png', 0.95):
                    _click_img('EmailNaoPodeSerEnviado.png', 0.9)
                    p.press('enter')
                    _wait_img('ChromeAberto.png', 0.9)
                    time.sleep(3)
                    p.hotkey('alt', 'f4')
                    time.sleep(1)
                    p.click(875, 142)
                    time.sleep(1)
                    _click_img('ExcluirEmail.png', 0.9)
                    _wait_img('ConfirmarExclusao.png', 0.9)
                    time.sleep(1)
                    p.press('s')
                    break

        _click_img('EmailEnviado.png', 0.9)
        p.press('enter')


@_time_execution
def run():
    envia_email()


if __name__ == '__main__':
    time.sleep(5)
    run()
