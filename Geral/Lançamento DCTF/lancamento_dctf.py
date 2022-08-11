from Dados import empresas
from datetime import datetime
import pyautogui as p
import time
import pyperclip


def esperar(imagen):
    while not p.locateOnScreen(imagen, confidence=0.9):
        time.sleep(1)


def clicar(imagen, tempo=0.1):
    try:
        p.click(p.locateCenterOnScreen(imagen, confidence=0.9))
        time.sleep(tempo)
        return True
    except:
        return False


def escrever(texto, tempo=0.1):
    try:
        with open('Lançamento DCTF.csv', 'a') as f:
            f.write(texto + '\n')
        time.sleep(tempo)
    except:
        with open('Lançamento DCTF - Auxiliar.csv', 'a') as f:
            f.write(texto + '\n')
        time.sleep(tempo)


def inserir_info():
    p.getWindowsWithTitle('DCTF Mensal')[0].maximize()
    prevEmpresa = ''
    for count, empresa in enumerate(empresas, start=1):
        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
        nome, cnpj, comp, ano, datAr, datVenc, datAp, codRe, valor = empresa
        print(' - '.join([indice, cnpj, nome]))

        meses = {'1': 'J', '2': 'F', '3': 'M', '4': 'A', '5': 'MM', '6': 'JJ',
                 '7': 'JJJ', '8': 'AA', '9': 'S', '10': 'O', '11': 'N', '12': 'D'}

        mes = meses[comp]

        if not prevEmpresa == nome:
            esperar(r'imagens\Nova.png')
            clicar(r'imagens\Nova.png')
            esperar(r'imagens\NovaDeclaracao.png')
            p.write(cnpj)
            time.sleep(0.5)
            p.press('tab')
            p.write(mes)
            time.sleep(0.5)
            p.press('tab')
            while not p.locateOnScreen(r'imagens\2021.png'):
                p.press('down')
            p.press('tab', presses=2)
            p.press('enter')

            while not p.locateOnScreen(r'imagens\DebitosCreditos.png'):
                if p.locateOnScreen(r'imagens\AbrirExistente.png'):
                    p.hotkey('alt', 's')
                if p.locateOnScreen(r'imagens\CNPJinvalido.png'):
                    p.press('enter')
                    escrever(';'.join([cnpj, nome, 'CNPJ invalido']))
                if p.locateOnScreen(r'imagens\RecuperarDados.png'):
                    p.hotkey('alt', 's')

            prevEmpresa = nome
        clicar(r'imagens\DebitosCreditos.png')
        time.sleep(0.5)


if __name__ == '__main__':
    inicio = datetime.now()
    try:

        inserir_info()

        print(datetime.now() - inicio)
    except:
        print(datetime.now() - inicio)
        p.alert(title='Script incrível', text='ERRO')
