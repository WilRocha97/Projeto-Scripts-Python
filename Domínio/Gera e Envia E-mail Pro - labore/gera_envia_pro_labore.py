# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, get_comp
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login


def gerar(comp, empresa, andamentos):
    cod, cnpj, nome = empresa
    p.hotkey('alt', 'r')
    time.sleep(1)
    p.press('r')
    time.sleep(1)
    p.press('f')
    time.sleep(1)
    while not _find_img('ReciboDePagamento.png', conf=0.9):
        time.sleep(1)
        
    for i in [0, 1]:
        p.write(comp)
        p.press('tab')

    time.sleep(1)
    p.press('f')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'o')
    
    cont = 1
    # espera o holerite gerar
    while not _find_img('HoleriteGerado.png', conf=0.9):
        time.sleep(1)
        if _find_img('impressora.png', conf=0.9):
            p.press('enter')
        if _find_img('SemDados.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Sem dados para gerar holerite')
            p.press('esc', presses=5, interval=0.3)
            print('❌ Sem dados para gerar holerite')
            return False
        if cont > 60:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Sem dados para gerar holerite')
            p.press('esc', presses=5, interval=0.3)
            print('❌ Sem dados para gerar holerite')
            return False
        cont += 1
    
    # clica no ícone do e-mail
    _click_img('email.png', conf=0.9)
    
    # espera o botão branco
    while not _find_img('EnviarPorEmail.png', conf=0.9):
        time.sleep(1)
    
    # clica no espaço da mensagem
    _click_img('Mensagem.png', conf=0.9)
    # guarda o texto
    pyperclip.copy('Esse e-mail foi enviado automaticamente, favor não responder.')
    # cola o texto
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('tab', presses=3, interval=0.5)
    p.press('enter')
    while not _find_img('Assinatura.png', conf=0.9):
        time.sleep(1)
    # escreve a assinatura
    if not _find_img('AssinaturaDoRobo.png', conf=0.9):
        p.write('At.te \nDepto Pessoal \nVeiga & Postal.')
    time.sleep(1)
    p.hotkey('alt', 'o')
    time.sleep(2)
    
    p.hotkey('alt', 'r')
    time.sleep(1)
    
    while _find_img('EnviarRelatorioPorEmail.png', conf=0.9):
        if _find_img('SemDestinatario.png', conf=0.9):
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Destinatário não informado')
            p.press('esc', presses=5, interval=0.3)
            print('❌ Destinatário não informado')
            return False
        time.sleep(1)
    
    _escreve_relatorio_csv(f'{cod};{cnpj};{nome};E-mail enviado com sucesso')
    p.press('esc', presses=5, interval=0.3)
    print('✔ E-mail enviado com sucesso')
    return True


@_time_execution
def run():
    comp = get_comp(printable='mm/yyyy', strptime='%m/%Y')
    if not comp:
        return False
    empresas = _open_lista_dados()
    andamentos = 'Gera e envia e-mail Pró - labore'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        
        if not _login(empresa, andamentos):
            continue
        gerar(comp, empresa, andamentos)


if __name__ == '__main__':
    run()
