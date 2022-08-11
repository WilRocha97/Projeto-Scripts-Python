# -*- coding: utf-8 -*-
from pyperclip import paste
from Dados import empresas
from datetime import date
from time import sleep
import pyautogui as a
import re, os, sys

sys.path.append('..')
from comum import find_img, click_img, wait_img, \
escrever_relatorio_csv, login_sn

def analisa_texto(texto, opcao):
    comps, ano_atual = [], str(date.today().year)
    regex = re.compile(r'\s+([A-Za-zç]+/\d{4})')
    lsaux = [(m.group(1), m.start(), m.end()) for m in regex.finditer(texto)]
    
    for i in range(len(lsaux)):
        comp, start = lsaux[i][0], lsaux[i][2]
        try:
            end = lsaux[i+1][1]
        except:
            end = len(texto)-1

        if 'R$' in texto[start:end]:
            if (not opcao and comp[-4:] == ano_atual): continue
            comps.append(comp)

    return comps


def consulta_debitos(opcao):
    print('>>> consultando debitos')
    exit, all_comps = False, []
    ano_atual = date.today().year

    a.press(('tab', 'tab', 'enter'))
    if not wait_img('select_ano.png', conf=0.9):
        return 'Timeout - Não carregou pagina consulta'

    sleep(0.2)
    while True:
        a.press('tab', presses=6)
        a.press('down', presses=2)
        a.press(('enter', 'tab', 'enter'))

        for i in range(30):
            if find_img('break.png', conf=0.9):
                exit = True
                break

            if find_img('txt_tabela.png', conf=0.9): break
            sleep(1)
        else:
            return 'Timeout - Não carregou a tabela de apuração'

        if exit: break
        a.hotkey('ctrl', 'a')
        a.hotkey('ctrl', 'c')
        comps = analisa_texto(paste(), opcao)
        if comps: all_comps.extend(comps)

    if all_comps:
        all_comps = 'debitos em ' + ', '.join(all_comps)
    else:
        all_comps = 'sem debitos'

    return all_comps


def run():
    opcao = a.confirm(text='Deseja incluir o ano atual na pesquisa?', buttons=['Sim', 'Não'])
    if opcao is None: return True
    opcao = opcao == 'Sim'

    total, lista_empresas = len(empresas), empresas
    for i, cnpj in enumerate(lista_empresas):
        print(i, total)

        status, text = login_sn(cnpj)
        if status: text = consulta_debitos(opcao)

        a.hotkey('ctrl', 'w')
        text = f'{cnpj};{text}'
        
        print(f'>>> {text}\n')
        escrever_relatorio_csv(text)

    return True


if __name__ == '__main__':
    if run():
        print('>>> Execução terminada com sucesso')
    else:
        print('>>> Execução terminada com erro')