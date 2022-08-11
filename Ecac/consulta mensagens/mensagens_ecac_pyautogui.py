# -*- coding: utf-8 -*-

import pyperclip, time, os, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def abrir():
    # Abrir o site
    if _find_img('Chrome.png', conf=0.9):
        _click_img('Chrome.png', conf=0.9)
    else:
        p.hotkey('win', 'm')
        time.sleep(0.5)
        _click_img('ChromeAtalho.png', conf=0.9, clicks=2)

    link = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    while not _find_img('link.png', conf=0.9, ):
        time.sleep(5)
        p.moveTo(1163, 377)
        p.click()
    _click_img('Maxi.png', conf=0.9, timeout=2)
    _click_img('link.png', conf=0.9)
    p.write(link)
    time.sleep(1)
    p.press('enter')
    time.sleep(3)

    if not _find_img('AlterarEmpresa.png', conf=0.9):
        _wait_img('GovBR.png', conf=0.9, timeout=-1)
        _click_img('GovBR.png', conf=0.9)

        # _wait_img('X.png', conf=0.9, timeout=-1)
        # _click_img('X.png', conf=0.9)

        _wait_img('Certificado.png', conf=0.9, timeout=-1)
        _click_img('Certificado.png', conf=0.9, )
        time.sleep(2)
        p.press('enter')
        _wait_img('AlterarEmpresa.png', conf=0.9, timeout=-1)


def login(index, empresas):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cert = empresa

        _indice(count, total_empresas, empresa)
        p.hotkey('win', 'm')

        if not _find_img('AlterarEmpresa.png', conf=0.9):
            abrir()

        while _find_img('Mensagens.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            abrir()

        _wait_img('AlterarEmpresa.png', conf=0.9, timeout=-1)
        _click_img('AlterarEmpresa.png', conf=0.9, )

        _wait_img('CNPJ.png', conf=0.9, timeout=-1)
        _click_img('CNPJ.png', conf=0.9)
        time.sleep(1)
        p.write(cnpj)
        time.sleep(1)

        _click_img('Alterar.png', conf=0.9)
        time.sleep(3)

        if _find_img('atencao.png', conf=0.9):
            _escreve_relatorio_csv('{};Erro no login'.format(cnpj))
            print('❌ Erro no login')
            continue

        if _find_img('Mensagens.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            _escreve_relatorio_csv('{};Mensagens no ECAC'.format(cnpj))
            print('❌ Mensagens no ECAC')
            continue

        _escreve_relatorio_csv('{};Não tem Mensagens no ECAC'.format(cnpj))
        print('Não tem Mensagens no ECAC')


@_time_execution
def run():
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    login(index, empresas)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
