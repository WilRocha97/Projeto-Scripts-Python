# -*- coding: utf-8 -*-
from sys import path
import requests

path.append(r'..\..\_comum')
from pyautogui_comum import get_comp
from comum_comum import download_file


def run():
    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp: return False

    comp = comp.split('/')
    comp.reverse()

    nome = f'TF' + ''.join(comp) + '.zip'

    url = f'http://www.quarta.com.br/downloads/gfip/{nome}'
    res = requests.get(url, verify=False)

    try:
        download_file(nome, res)
        print(f'Indice {nome} baixado')
    except Exception as e:
        print(f'Não pode baixar o indice devido {str(e)}')


if __name__ == '__main__':
    run()