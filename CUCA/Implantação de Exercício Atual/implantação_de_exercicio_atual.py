# -*- coding: utf-8 -*-
from datetime import datetime
from sys import path
import pyautogui as p
import time

path.append(r'..\..\_comum')
from cuca_comum import _horario, _fechar, _login, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def verificacoes(empresa, execucao):
    cod, cnpj, valor, comp, ano = empresa
    if _find_img('Invalido.png', conf=0.9):
        _click_img('Invalido.png', conf=0.9)
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj, 'Opção inválida para esta empresa!']), nome=execucao)
        print('** Opção inválida para esta empresa! **')
        return False
    if _find_img('Exclusiva.png', conf=0.9):
        _click_img('Exclusiva.png', conf=0.9)
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj, 'Opção exclusiva da Matriz!']), nome=execucao)
        print('** Opção exclusiva da Matriz! **')
        return False
    if _find_img('InvalidoMEI.png', conf=0.9):
        _click_img('InvalidoMEI.png', conf=0.9)
        p.press('enter')
        p.press(['esc'], presses=2, interval=1)
        _escreve_relatorio_csv(';'.join([cnpj, 'Opção inválida para empresa optante pelo MEI']), nome=execucao)
        print('** Opção inválida para esta empresa MEI **')
        return False
    if _find_img('BalancoFechado.png', conf=0.9):
        _click_img('BalancoFechado.png', conf=0.9)
        p.press('enter')
        p.press(['esc'], presses=2, interval=1)
        _escreve_relatorio_csv(';'.join([cnpj, 'Balanço fechado no mês ' + comp]), nome=execucao)
        print('** Balanço fechado no mês ' + comp + ' **')
        return False
    if _find_img('SoMatriz.png', conf=0.9):
        _click_img('SoMatriz.png', conf=0.9)
        p.press('enter')
        p.press(['esc'], presses=2, interval=1)
        _escreve_relatorio_csv(';'.join([cnpj, 'Alterações somente na Matriz']), nome=execucao)
        print('** Alterações somente na Matriz **')
        return False
    return True


def implantacao(empresas, execucao, index):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'CUCA'):
            _iniciar('cuca')
            p.getWindowsWithTitle('(SP)')[0].maximize()

        _indice(count, total_empresas, empresa, index)

        cod, cnpj, valor, comp, ano = empresa

        # Defini os textos usados nas respostas das verificações de login e empresa
        if not _login(empresa, 'CNPJ', 'cuca', 'Implantação de Exercício Atual', comp, ano):
            continue

        texto = '{};Empresa não encontrada no CUCA'.format(cnpj)
        if not _verificar_empresa(cnpj, execucao, texto, 'cuca'):
            continue

        while not _find_img('Implantacao.png', conf=0.9,):
            # Abrir tela de Implantação
            p.hotkey('alt', 'l')
            time.sleep(0.5)
            p.press('a')
            time.sleep(0.5)
            p.press('up', presses=3)
            time.sleep(0.5)
            p.press('enter')

        if not verificacoes(empresa, execucao):
            continue

        time.sleep(1)
        p.hotkey('ctrl', 'i')
        _wait_img('Despesas.png', conf=0.9, timeout=-1)

        if not verificacoes(empresa, execucao):
            continue

        _click_img('Despesas.png', conf=0.9)
        _wait_img('Despesas2.png', conf=0.9, timeout=-1)
        p.press('enter')

        if not verificacoes(empresa, execucao):
            continue

        _wait_img('Campo.png', conf=0.9, timeout=-1)
        if comp == '1':
            _click_img('Valor.png', conf=0.9, clicks=2)
        _wait_img('Campo.png', conf=0.9, timeout=-1)
        time.sleep(0.5)
        p.write(valor)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.press(['esc'], presses=2, interval=1)
        _escreve_relatorio_csv(';'.join([cnpj, 'Implantação concluída']), nome=execucao)
        print('✔ Implantação concluída')

        _inicial('cuca')


@_time_execution
def run():
    execucao = 'Implantação de Exercício Atual'
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('CUCA')
    implantacao(empresas, execucao, index)
    _fechar()


if __name__ == '__main__':
    run()
