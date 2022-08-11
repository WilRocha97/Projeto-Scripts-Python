# -*- coding: utf-8 -*-
from datetime import datetime
from sys import path
import pyperclip, time, pyautogui as p

path.append(r'..\..\_comum')
from cuca_comum import _horario, _fechar, _login, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def login(empresa, execucao):
    cod, cnpj, nome = empresa
    texto = '{};{};{};Empresa não encontrada no CUCA'.format(cod, cnpj, nome)
    if not _login(empresa, 'Codigo', 'cuca', execucao):
        return False
    if not _verificar_empresa(cnpj, execucao, texto, 'cuca'):
        return False
    return True


def consulta_faturamento_vs_compra(execucao, empresas, index):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)

        # Verificar horário
        if _horario(_hora_limite, 'CUCA'):
            _iniciar('cuca')
            p.getWindowsWithTitle('(SP)')[0].maximize()

        _indice(count, total_empresas, empresa)

        cod, cnpj, nome = empresa

        _inicial('CUCA')

        if not login(empresa, execucao):
            continue

        # Abrir tela de Faturamento
        p.hotkey('alt', 'l')
        time.sleep(0.5)
        p.press(['e',  'd',  'enter'])
        time.sleep(1)

        if _find_img('Impossivel.png', 0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Impossível  continuar, balanço não foi fechado corretamente.']), nome=execucao)
            continue

        while not _find_img('DeclaracaoDeFaturamento.png', 0.9):
            # As vezes aparece uma tela para iniciar a tabela do mês da empresa, deve confirmar
            if _find_img('Continuar.png', 0.9):
                p.hotkey('alt', 's')
                time.sleep(10)
            if not _find_img('DeclaracaoDeFaturamento.png', 0.9):
                # Se a tela de faturamento não abrir, tenta de novo
                p.hotkey('alt', 'l')
                time.sleep(1)
                # Se aparecer cancela a nova tentativa
                if _find_img('DeclaracaoDeFaturamento.png', 0.9):
                    break
                p.press(['e', 'd', 'enter'])
                time.sleep(3)

        # Enquanto estiver selecionado a coluna do meio ele tenta sair da seleção
        while _find_img('Fatura.png'):
            p.press('esc')
            time.sleep(1)
        time.sleep(1)
        # Aperte TAB 23 vezes para selecionar o primeiro campo da coluna de totais para começar a copiar as informações
        if _find_img('Jan.png', conf=0.9):
            p.press('tab', presses=23)
        time.sleep(1)

        # Copiar e colar os valores totais de cada mês de faturamento e compra
        for campo in ('Faturamento', 'Compra'):
            cod, cnpj, nome = empresa
            dados = ';'.join([cod, cnpj, nome, campo])
            for i in range(12):
                p.press('tab')
                time.sleep(0.1)
                p.hotkey('ctrl', 'c')
                campo = pyperclip.paste()
                dados += ";" + campo

            _escreve_relatorio_csv(dados, nome=execucao)

        print('✔ Consulta concluída')
        _inicial('CUCA')


@_time_execution
def run():
    execucao = 'Faturamento VS Compra 3'
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('CUCA')
    consulta_faturamento_vs_compra(execucao, empresas, index)
    _fechar()


if __name__ == '__main__':
    run()
