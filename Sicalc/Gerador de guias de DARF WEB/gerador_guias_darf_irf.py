# -*- coding: utf-8 -*-
import os, time, shutil, re, pyperclip, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice, _headers, e_dir
from pyautogui_comum import _click_img, _wait_img, _find_img, _click_position_img


def login_sicalc(empresa):
    cnpj, nome, nota, valor, cod = empresa
    p.hotkey('win', 'm')

    # Abrir o site
    if _find_img('Chrome.png', conf=0.99):
        pass
    elif _find_img('ChromeAberto.png', conf=0.99):
        _click_img('ChromeAberto.png', conf=0.99, timeout=1)
    else:
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        while not _find_img('Google.png', conf=0.9):
            time.sleep(1)
            p.moveTo(1163, 377)
            p.click()
    
    """while not _find_img('remover_dados.png', conf=0.9):
        p.hotkey('ctrl', 'shift', 'del')
        time.sleep(1)
        
    p.press('tab')
    p.press('enter')
    while _find_img('remover_dados.png', conf=0.9):
        time.sleep(1)
    p.hotkey('ctrl', 'w')
    time.sleep(1)"""
    
    link = 'https://sicalc.receita.economia.gov.br/sicalc/rapido/contribuinte'
    
    _click_img('Maxi.png', conf=0.9, timeout=1)
    p.click(1100, 51)
    time.sleep(1)
    p.write(link.lower())
    time.sleep(1)
    p.press('enter')
    time.sleep(3)
    
    # selecionar o checkbox do captcha
    p.press('pgDn')
    _click_img('checkbox.png', conf=0.95)
    p.moveTo(5, 5)
    
    timer = 0
    while not _find_img('check.png', conf=0.9):
        if _find_img('erro_login.png', conf=0.9):
            return False
        if _find_img('site_morreu.png', conf=0.9):
            return False
        if _find_img('site_bugou.png', conf=0.9):
            return False
        if _find_img('site_bugou_2.png', conf=0.9):
            return False
        if _find_img('site_bugou_3.png', conf=0.9):
            return False
        time.sleep(1)
        timer += 1
        if timer >= 15:
            return False
    
    _click_img('possoa_juridica.png', conf=0.95)
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    p.write(cnpj)
    time.sleep(1)
    p.press('tab', presses=5, interval=0.1)
    time.sleep(1)
    
    # esperar o menu referente a guia aparecer
    timer = 0
    while not _find_img('menu.png', conf=0.9):
        if _find_img('erro.png', conf=0.9):
            return False
        time.sleep(1)
        p.press('enter')
        timer += 1
        if timer >= 10:
            return False
    
    return True
    

def gerar(empresa, apuracao):
    cnpj, nome, nota, valor, cod = empresa
    print('>>> Inserindo informações da guia')
    
    # insere observação
    nota.split(',')
    p.write(f'NF: {nota}')
    time.sleep(1)
    
    # inserir o código da receita referente a guia
    p.press('pgDn')
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    p.write(cod)
    time.sleep(1)
    
    # confirmar a seleção
    p.press(['down', 'enter'], interval=1)
    time.sleep(1)
    p.moveTo(106, 375)
    p.click()
    time.sleep(1)

    # descer a visualização da página
    timer = 0
    while not _find_img('apuracao.png', conf=0.9):
        if _find_img('erro.png', conf=0.9):
            return False, ''
        p.press('pgDn')
        time.sleep(1)
        timer += 1
        if timer >= 15:
            return False, ''
        
    # clicar no campo e inserir a data de apuração
    _click_img('apuracao.png', conf=0.9)
    time.sleep(1)
    p.write(apuracao)
    time.sleep(1)
    p.press('tab', presses=2, interval=1)
    time.sleep(1)
    p.hotkey('ctrl', 'c')
    p.hotkey('ctrl', 'c')
    vencimento = pyperclip.paste()
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    p.write(valor)

    p.press('tab')
    time.sleep(1)

    # espera guia ser calculada
    timer = 0
    while not _find_img('checkbox_guia.png', conf=0.9):
        if _find_img('erro.png', conf=0.9):
            return False, ''
        _click_img('calcular.png', conf=0.95, timeout=1)
        # descer a visualização da página
        p.press('pgDn')
        time.sleep(1)
        timer += 1
        if timer >= 15:
            return False, ''
        
    print('>>> Guia calculada')
    
    # marca a guia que será gerada
    _click_position_img('checkbox_guia.png', '+', pixels_y=12, conf=0.95)
    time.sleep(1)
    
    print('>>> Gerando guia')
    # _click_img('emitir_darf.png', conf=0.95)
    return True, vencimento


def salvar_guia(empresa, apuracao, vencimento, tipo):
    cnpj, nome, nota, valor, cod = empresa
    nome = nome[:10]
    nome = nome.replace(' PAU ', ' P ').replace(' CU ', ' C ').replace(' CUS ', ' C ').replace(' CUM ', ' C ')
    
    timer = 0
    while not _find_img('SalvarComo.png', conf=0.9):
        _click_img('emitir_darf.png', conf=0.9)
        time.sleep(2)
        if _find_img('site_morreu_salvar.png', conf=0.9):
            p.hotkey('ctrl', 'w')
        timer += 2
        if timer >= 15:
            return False
            
    # exemplo: NOME DA EMPRESA - 00000000000 - DARF IRRF 170806 02-2023 - venc. 20-03-2023.pdf
    p.write(f'{nome.replace("/", " ")} - {cnpj} - {tipo} {cod} {apuracao.replace("/", "-")} - venc. {vencimento.replace("/", "-")}.pdf')
    time.sleep(0.5)
    
    pasta_final = r'\\vpsrv03\Arq_Robo\Gerador de guias de DARF IRRF\Ref. ' + apuracao.replace("/", "-") + '\Guias'
    os.makedirs(pasta_final, exist_ok=True)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    pyperclip.copy(pasta_final)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('SalvarComo.png', conf=0.9):
        if _find_img('Substituir.png', conf=0.9):
            p.press('s')
    time.sleep(2)
    p.hotkey('ctrl', 'w')
    return True

    
@_time_execution
def run():
    # p.mouseInfo()
    tipo = p.confirm(title='Script incrível', text='Qual tipo da guia?', buttons=('DARF IRRF', 'DARF DP'))
    apuracao = p.prompt(title='Script incrível', text='Qual o período de apuração?', default='00/0000')
    
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    p.hotkey('win', 'm')
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome, nota, valor, cod = empresa
        # configurar o indice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        
        erro = 'sim'
        while erro == 'sim':
            # try:
            # fazer login do SICALC
            if not login_sicalc(empresa):
                p.hotkey('ctrl', 'w')
                erro = 'sim'
            else:
                # gerar a guia de DCTF
                resultado, vencimento = gerar(empresa, str(apuracao))
                if not resultado:
                    p.hotkey('ctrl', 'w')
                    erro = 'sim'
                else:
                    if not salvar_guia(empresa, apuracao, vencimento, tipo):
                        erro = 'sim'
                    else:
                        print('✔ Guia gerada')
                        _escreve_relatorio_csv('{};{};{};{};Guia gerada'.format(cnpj, nome, valor, cod), nome=f'Resumo gerador {tipo}')
                        erro = 'nao'
            """except:
                try:
                    p.hotkey('alt', 'f4')
                    erro = 'sim'
                except:
                    erro = 'sim'"""

        p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
