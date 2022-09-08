# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path
from pygetwindow import getWindowsWithTitle

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def importar(empresa):
    empresa = empresa[0]
    _wait_img('SEDIF.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.hotkey('alt', 'i')
    p.hotkey('alt', 'i')
    
    _wait_img('ImportarArquivo.png', conf=0.9, timeout=-1)
    time.sleep(1)

    _click_img('ImportarArquivo.png', conf=0.9)
    
    _click_img('Abrir.png', conf=0.9, clicks=2)
    
    _wait_img('AbrirArquivo.png', conf=0.9, timeout=-1)
    
    p.write(empresa)
    p.press('enter')
    
    time.sleep(4)
    
    p.hotkey('alt', 'i')
    
    while not _find_img('Importado.png', conf=0.9):
        if _find_img('JaExiste.png', conf=0.9):
            _escreve_relatorio_csv(f'{empresa};Já importado', 'Importação SEDIF')
            print('❗ Arquivo já importado')
            p.hotkey('alt', 'n')
            time.sleep(2)
            if _find_img('Atencao.png', conf=0.9):
                p.press('enter')
            _wait_img('Importado.png', conf=0.9, timeout=-1)
            p.press('enter')
            return False
            
        if _find_img('ErroNoArquivo.png', conf=0.9):
            _escreve_relatorio_csv(f'{empresa};Arquivo com erros', 'Importação SEDIF')
            print('❌ Arquivo com erros')
            p.press('enter')
            time.sleep(2)
            if _find_img('ErroValidacao.png', conf=0.9):
                p.hotkey('alt', 'f')
            return False

    _escreve_relatorio_csv(f'{empresa};Arquivo importado', 'Importação SEDIF')
    print('✔ Arquivo importado')
    p.press('enter')
    

@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        importar(empresa)


if __name__ == '__main__':
    run()
