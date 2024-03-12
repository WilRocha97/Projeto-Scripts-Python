# -*- coding: utf-8 -*-
import pyperclip, time, os, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login, _salvar_pdf, _login_web, _abrir_modulo


def verifica(empresa, andamento):
    cod, cnpj, nome = empresa
    
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'c')
    time.sleep(0.5)
    # Impostos
    p.press('p')
    time.sleep(0.5)
    
    while not _find_img('personaliza.png', conf=0.9):
        time.sleep(1)
    _click_img('personaliza.png', conf=0.9)

    while not _find_img('aviso_previo.png', conf=0.9):
        time.sleep(1)
    _click_img('aviso_previo.png', conf=0.9)
    
    while not _find_img('opcao.png', conf=0.9):
        time.sleep(1)

    if _find_img('opcao_marcada.png', conf=0.99):
        _click_img('opcao_marcada.png', conf=0.99)
        time.sleep(0.5)
        p.hotkey('alt', 'g')
        while _find_img('tela_aberta.png', conf=0.9):
            if _find_img('existem_calculos.png', conf=0.95):
                _click_img('existem_calculos.png', conf=0.95)
                p.hotkey('alt', 'y')
            if _find_img('aviso_1.png', conf=0.95):
                p.hotkey('alt', 'o')
                time.sleep(1)
                p.press('esc', presses=4)
                _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Não é possível desmarcar a opção;Informe apenas uma empresa como centralizadora ou configure todas as '
                                       f'empresas como centralizada indicando que a centralizadora consta em outro banco de dados.', nome=andamento)
                print('❌ Não é possível desmarcar a opção')
                return
            if _find_img('aviso_2.png', conf=0.95):
                p.hotkey('alt', 'o')
                time.sleep(1)
                p.press('esc', presses=4)
                _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Não é possível desmarcar a opção;É necessário configurar o grupo "Configuração do relatório de provisão das '
                                       f'férias calculadas no mês".', nome=andamento)
                print('❌ Não é possível desmarcar a opção')
                return
            if _find_img('aviso_3.png', conf=0.9):
                _click_img('aviso_3.png', conf=0.9)
                time.sleep(1)
                p.hotkey('alt', 'o')
                time.sleep(1)
                _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Opção desmarcada;Deverá ser informado ao eSocial o processo de limitação do cálculo de terceiros até 20 salários '
                                       f'mínimos através do evento S-1070 e os códigos de terceiros através do cadastro do serviço, ta tela "Cód. Terceiros Suspenso".', nome=andamento)
                print('❌ Não é possível desmarcar a opção')
                return
            if _find_img('aviso_4.png', conf=0.9):
                p.hotkey('alt', 'o')
                time.sleep(1)
                p.press('esc', presses=4)
                _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Não é possível desmarcar a opção;A empresa informada já iniciou o envio das informações em outro banco de dados. Para que seja '
                                       f'possível continuar o envio das informações nesse banco de dados, devera ser efetuada a Importação de Dados do eSocial oi a Conversão '
                                       f'de Dados do eSocial.', nome=andamento)
                print('❌ Não é possível desmarcar a opção')
                return
            if _find_img('aviso_5.png', conf=0.9):
                p.hotkey('alt', 'n')
                time.sleep(1)
                p.press('esc', presses=4)
                _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Não é possível desmarcar a opção;A tela de "Apuração previdenciária" somente será emitida pelo módulo Folha, '
                                       f'pois na Escrita a mesma apenas é emitida para empresas consideradas matriz.', nome=andamento)
                print('❌ Não é possível desmarcar a opção')
                return
            time.sleep(1)

        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Opção desmarcada', nome=andamento)
        print('❗ Opção desmarcada')
    else:
        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Já está desmarcada', nome=andamento)
        print('✔ Já esta desmarcada')
    
    p.press('esc', presses=4)
    time.sleep(2)
    
    
@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('folha')
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)
    
        if not _login(empresa, andamentos):
            continue
            
        verifica(empresa, andamentos)


if __name__ == '__main__':
    empresas = _open_lista_dados()
    andamentos = 'Verifica Aviso Prévio'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
