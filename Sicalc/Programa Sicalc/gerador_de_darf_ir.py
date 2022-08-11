from dados import empresas as emp
from datetime import datetime
from sys import path
import pyautogui as p
import time
import pyperclip
import os

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from cuca_comum import _escrever


def abrir_sicalc():
    if p.getWindowsWithTitle('Sicalc Auto Atendimento'):
        if not p.getWindowsWithTitle('Sicalc Auto Atendimento')[0].isMaximized:
            p.getWindowsWithTitle('Sicalc Auto Atendimento')[0].maximize()
            _wait_img('TelaInicialLancamento.png', conf=0.9, timeout=-1)
    else:
        _click_img('Sicalc.png', conf=0.9, clicks=2)
        _wait_img('Esclarecimento.png', conf=0.9, timeout=-1)
        _click_img('Esclarecimento.png', conf=0.9)
        p.hotkey('alt', 'c')
        time.sleep(1)
        p.hotkey('alt', 'c')
        _wait_img('TelaInicialLancamento.png', conf=0.9, timeout=-1)


def fechar_sicalc():
    _click_img('TelaInicialLancamento.png', conf=0.9)
    p.hotkey('alt', 'p')
    _wait_img('Atencao.png', conf=0.9, timeout=-1)
    p.hotkey('alt', 's')
    time.sleep(1.5)
    p.getWindowsWithTitle('Sicalc Auto Atendimento')[0].close()


def gerar_guia():
    for count, empresa in enumerate(emp, start=1):
        if count != 1:
            print('[ {} Restantes ]\n'.format((len(emp)) - (count - 1)))

        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(emp)) + ' ]'
        nome, cnpj, valor, cod = empresa

        # Focar no aplicativo
        p.getWindowsWithTitle('Sicalc Auto Atendimento')[0].maximize()
        time.sleep(0.5)
        _click_img('TelaInicialLancamento.png', conf=0.9)

        # Limpar dados antigos
        p.hotkey('alt', 'p')
        _wait_img('Atencao.png', conf=0.9, timeout=-1)
        p.hotkey('alt', 's')
        time.sleep(1.5)
        _click_img('CodigoReceita.png', conf=0.9)

        print(' - '.join([indice, cod, nome, cnpj]))
        # Navegar e preencher os campos
        p.write(cod)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press(['m', 'tab'], interval=0.5)
        time.sleep(0.5)
        p.write(comp + ano2)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write(valor)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)

        # Calcular guia
        p.hotkey('alt', 'l')
        time.sleep(0.5)

        # Preencher dados da empresa
        p.hotkey('alt', 'f')
        _wait_img('Preenchimento.png', conf=0.9, timeout=-1)
        p.write(nome)
        time.sleep(0.5)
        p.press(['tab'], presses=2, interval=0.5)
        time.sleep(0.5)
        p.write(cnpj)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(1)
        if _find_img('Invalido.png', conf=0.9):
            p.press('enter')
            time.sleep(0.5)
            invalido = ';'.join([nome, cnpj, cod, 'CPF/CNPJ Inválido'])
            _escrever('Gerador de DARF IRF', invalido)
            print('** CPF/CNPJ Inválido **\n\n')
            _click_img('Fechar.png', conf=0.9)
            continue
        if _find_img('SoMatriz.png', conf=0.9):
            p.hotkey('alt', 's')
        time.sleep(1)

        # Gerar e salvar o PDF da guia
        p.hotkey('alt', 'i')
        _wait_img('SalvarComo.png', conf=0.9, timeout=-1)
        time.sleep(1)
        # Usa o pyperclip porque o pyautogui não digita letra com acento
        '''pyperclip.copy(';'.join([cnpj, 'IRRF', 'Nota Fiscal', mes, ano, 'Padrão', venc, nome]) + ' - ' + cod + '.pdf')
        p.hotkey('ctrl', 'v')'''
        p.write(cnpj + ' - ' + cod + '.pdf')
        time.sleep(1)

        p.press('tab', presses=6)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        pyperclip.copy('V:\Setor Robô\Scripts Python\Sicalc\Programa Sicalc\Guias')
        p.hotkey('ctrl', 'v')
        time.sleep(0.5)
        p.press('enter')

        time.sleep(0.5)
        p.hotkey('alt', 'l')
        time.sleep(0.5)
        p.hotkey('alt', 'c')

        # Anotar na planilha de Andamentos
        _escrever('Gerador de DARF IRF', ';'.join([nome, cnpj, cod, 'Guia gerada']))
        print('>>> Guia gerada <<<')


if __name__ == '__main__':
    os.makedirs('Guias', exist_ok=True)
    # Pergunta a data de vencimento1
    venc = p.prompt(title='Script incrível', text='Qual o vencimento?', default='-')

    # Define a competência que ira digitar no aplicativo para gerar a guia
    comp = datetime.now().month
    # Se o script rodar no mês 1 a competencia da guia é 12 e o ano também tem que ser atualizado para o ano anterior
    if comp == 1:
        comp = 12
        ano = datetime.now().year
        ano -= 1
        year = datetime.now().replace(year=ano)
    else:
        comp -= 1
        year = datetime.now()

    # Pega a competencia atualizada para a guia já onvertida em string
    month = datetime.now().replace(month=comp)
    comp = month.strftime('%m')

    # Define o ano que será digitado no nome do arquivo e o ano que sera digitado no aplicativo para gerar a guia já convertida em string
    ano = year.strftime('%Y')
    ano2 = year.strftime('%y')

    # Define o mês que seja digitado no nome do arquivo
    mes = {'01': '01 - Janeiro', '02': '02 - Fevereiro', '03': '03 - Março', '04': '04 - Abril', '05': '05 - Maio',
           '06': '06 - Junho', '07': '07 - julho', '08': '08 - Agosto', '09': '09 - Setembro', '10': '10 - Outubro',
           '11': '11 - Novembro', '12': '12 - Dezembro'}
    mes = mes[comp]

    p.hotkey('win', 'm')
    time.sleep(0.5)
    abrir_sicalc()
    gerar_guia()
    fechar_sicalc()
