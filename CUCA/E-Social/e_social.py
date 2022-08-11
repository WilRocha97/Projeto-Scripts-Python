# -*- coding: utf-8 -*-
from datetime import datetime
from sys import path
import pyautogui as p
import time

path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def enviar(execucao, empresa, mes, ano):
    cod, cnpj, nome = empresa
    while not _find_img('Minimizar.png', conf=0.9):
        if _find_img('DesejaGerar.png', conf=0.9):
            p.hotkey('alt', 's')
        if _find_img('Mensagem.png', conf=0.9):
            if _find_img('NaoGerouCarga.png', conf=0.9):
                p.hotkey('alt', 'o')
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'Não existem eventos pendentes ou não possui movimentação na competência']), nome=execucao)
                print('❗ Não existem eventos pendentes ou não possui movimentação na competência')
                p.press('esc')
                return False
            p.hotkey('alt', 'o')
            time.sleep(2)

    # Minimizar e selecionar todos os itens
    if not _find_img('Minimizar.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'Não gerou arquivos']), nome=execucao)
        print('❌ Não gerou arquivos')
        return False
    while _find_img('Minimizar.png', conf=0.9):
        while _find_img('Minimizar.png', conf=0.9):
            _click_img('Minimizar.png', conf=0.9)

        while _find_img('Check.png', conf=0.9):
            _click_img('Check.png', conf=0.9)

    # Enviar os itens selecionados
    # p.hotkey('alt', 'n')
    time.sleep(1)

    # Esperar a tela de aguardar fechar
    while _find_img('Aguarde.png', conf=0.9):
        time.sleep(2)
        while _find_img('Aguarde.png', conf=0.9):
            time.sleep(2)
            while _find_img('Aguarde.png', conf=0.9):
                time.sleep(2)
                while _find_img('Aguarde.png', conf=0.9):
                    time.sleep(2)
                    while _find_img('Aguarde.png', conf=0.9):
                        time.sleep(2)

    return True


def social(empresas, mes, ano, index, execucao):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        '''_hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('DPCUCA')
            p.getWindowsWithTitle('DPCUCA')[0].maximize()'''

        _indice(count, total_empresas, empresa)

        cod, cnpj, nome = empresa

        # Verificações de login e empresa
        if not _login(empresa, 'Codigo', 'dpcuca', execucao, mes, ano):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'Empresa Inativa']), nome=execucao)
            print('❌ Empresa Inativa')
            continue

        # CNPJ com os separadores para poder verificar a empresa no cuca
        texto = '{};{};{};{};{};Robô sem acesso no DPCUCA'.format(cod, cnpj, nome, mes, ano)
        if not _verificar_empresa(cnpj, execucao, texto, 'dpcuca'):
            continue

        if _find_img('ferias.png', conf=0.9):
            _click_img('ferias.png', conf=0.9)
            _click_img('feriasFechar.png', conf=0.9)

        # Entrar no e-social
        while not _find_img('OK.png', conf=0.9):
            # Fechar a janela de férias que impede a exclusão dos itens
            if _find_img('ferias.png', conf=0.9):
                _click_img('ferias.png', conf=0.9)
                _click_img('feriasFechar.png', conf=0.9)
            _click_img('eSocial.png', conf=0.9)
            _click_img('eSocial2.png', conf=0.9)
            time.sleep(2)

        # Fechar a mensagem que aparece
        time.sleep(1)
        if _find_img('OK.png', conf=0.9):
            _click_img('OK.png', conf=0.9)

        if _find_img('Minimizar.png', conf=0.9):
            # Minimizar e selecionar todos os itens
            while _find_img('Minimizar.png', conf=0.9):
                # Fechar a janela de férias que impede a exclusão dos itens
                if _find_img('ferias.png', conf=0.9):
                    _click_img('ferias.png', conf=0.9)
                    _click_img('feriasFechar.png', conf=0.9)
                while _find_img('Minimizar.png', conf=0.9):
                    _click_img('Minimizar.png', conf=0.9)

                while _find_img('Check.png', conf=0.9):
                    _click_img('Check.png', conf=0.9)

            # Fechar a janela de férias que impede a exclusão dos itens
            if _find_img('ferias.png', conf=0.9):
                _click_img('ferias.png', conf=0.9)
                _click_img('feriasFechar.png', conf=0.9)

            # Excluir os itens
            while not _find_img('TemCerteza.png', conf=0.9):
                _click_img('Excluir.png', conf=0.9)
            p.hotkey('alt', 's')

        _wait_img('Limpo.png', conf=0.9)

        # Gerar Fechamento
        p.click(630, 206)

        p.hotkey('ctrl', 'n')
        _wait_img('Cargas.png', conf=0.9)

        while not _find_img('CargaFechamento2.png', conf=0.9):
            _click_img('CargaFechamento.png', conf=0.9)

        p.hotkey('ctrl', 'n')

        while not _find_img('EscolherEmpresa.png', conf=0.9):
            if _find_img('Responsavel.png', conf=0.9):
                p.hotkey('alt', 'g')

        p.hotkey('alt', 'a')
        time.sleep(2)

        if not enviar(execucao, empresa, mes, ano):
            p.press('esc')
            continue

        time.sleep(2)
        _click_img('Busca.png', conf=0.9)
        p.write('1299')
        time.sleep(2)
        if _find_img('SemErros.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'OK']), nome=execucao)
            print('✔ OK')
        else:
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, '1299']), nome=execucao)
            print('❌ 1299')
        p.press('esc')
        _inicial('DPCUCA')


@_time_execution
def run():
    execucao = 'E-Social'
    empresas = _open_lista_dados()

    data = p.prompt(text='Qual a competência?', title='Script incrível', default='00/0000')
    data = data.split('/')
    mes, ano = data[0], data[1]
    if int(mes) < 10:
        mes = str(mes)
        mes = mes.replace('0', '')

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('DPCUCA')
    social(empresas, str(mes), str(ano), index, execucao)
    _fechar()


if __name__ == '__main__':
    run()
