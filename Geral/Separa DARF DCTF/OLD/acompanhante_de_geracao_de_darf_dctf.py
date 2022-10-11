import pyautogui as p
import time
import pyperclip


def imprimir():
    # Esperar o PDF gerar
    time.sleep(1)
    while not p.locateOnScreen(r'imagens\PDF.png', confidence=0.9):
        time.sleep(0.5)

    # Salvar o PDF
    while not p.locateOnScreen(r'imagens\SalvarPDF.png', confidence=0.9):
        p.click(p.locateCenterOnScreen(r'imagens\PDF.png'))
    time.sleep(1)

    # Usa o pyperclip porque o pyautogui não digita letra com acento
    pyperclip.copy('Relatorios')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Geral\Separa DARF DCTF\PDFs')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')

    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if p.locateCenterOnScreen(r'imagens\SalvarComo.png', confidence=0.9):
        p.press('s')
        time.sleep(1)
    while p.locateOnScreen(r'imagens\Progresso.png', confidence=0.9):
        time.sleep(0.5)
    while p.locateOnScreen(r'imagens\PDF.png', confidence=0.9):
        p.press('esc')
        time.sleep(1)
        if p.locateOnScreen(r'imagens\Comunicado.png', confidence=0.9):
            p.press(['right', 'enter'], interval=0.2)
            time.sleep(1)
    return True


while not p.locateOnScreen(r'imagens\gerou.png'):
    if p.locateOnScreen(r'imagens\criticas.png'):
        p.press('esc')
    if p.locateOnScreen(r'imagens\alerta.png'):
        p.press('enter')

