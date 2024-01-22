# -*- coding: utf-8 -*-
import time, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def calculos():
    # Entrar na opção de Cálculos
    time.sleep(2)
    while not _find_img('SelecionarEvento.png', conf=0.9):
        _click_img('CUCAInicial.png', conf=0.9)
        p.hotkey('alt', 'c')
        time.sleep(0.5)

    evento, imagen, cont = 'sócio', 'Socios', 3

    p.press('e', presses=cont, interval=0.2)
    p.press('enter')
    _wait_img('Eventos{}.png'.format(imagen), conf=0.9, timeout=-1)
    time.sleep(3)
    return evento, imagen


def adicionar(andamentos, empresa):
    cod, cnpj, nome, mes, ano, cod_socio, nome_socio, data, valor = empresa

    if not _find_img('Tabela.png', conf=0.9):
        p.hotkey('alt', 'l')

    _wait_img('Tabela.png', conf=0.9, timeout=-1)

    if _find_img('PorCodigo.png', conf=0.99):
        _click_img('PorCodigo.png', conf=0.99)

    _click_img('FiltroCodigo.png', conf=0.99)
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.write(cod_socio)
    time.sleep(0.5)

    p.hotkey('alt', 'i')

    _wait_img('NovoRegistro.png', conf=0.9, timeout=-1)
    time.sleep(0.5)

    p.write(data)
    time.sleep(0.5)

    p.press('tab', presses=2)
    time.sleep(0.5)

    p.write(valor)
    time.sleep(0.5)

    p.hotkey('alt', 'g')

    if _find_img('JaExisteRegistro.png', conf=0.9):
        _escreve_relatorio_csv('{};{};{};{};{};{};Já existe registro'.format(cod, cnpj, nome, nome_socio, data, valor), nome=andamentos)
        print('✔ Já existe registro')
        p.press('enter')
        time.sleep(0.5)
        p.press('esc')
        time.sleep(0.5)
        p.press('N')

    else:
        _escreve_relatorio_csv('{};{};{};{};{};{};Inclusão realizada'.format(cod, cnpj, nome, nome_socio, data, valor), nome=andamentos)
        print('✔ Inclusão realizada')


def entrar_na_empresa(index, empresas, andamentos, usuario):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod, cnpj, nome, mes, ano, cod_socio, nome_socio, data, valor = empresa

        _indice(count, total_empresas, empresa, index)

        # CNPJ com os separadores para poder verificar a empresa no cuca
        if not _verificar_empresa(cnpj, andamentos, texto='falso', qual_cuca='dpcuca'):
            response = 'continue'
        else:
            response = 'mesma empresa'

        if response == 'continue':
            '''_hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
            # Verificar horário
            if _horario(_hora_limite, 'DPCUCA'):
                _iniciar('dpcuca')
                p.getWindowsWithTitle('DPCUCA')[0].maximize()'''

            # Verificações de login
            if not _login(empresa, 'Codigo', 'dpcuca', andamentos, mes, ano):
                p.press('enter')
                _escreve_relatorio_csv('{};{};{};{};{};{};Empresa Inativa'.format(cod, cnpj, nome, nome_socio, data, valor), nome=andamentos)
                print('❌ Empresa Inativa')
                continue

            # CNPJ com os separadores para poder verificar a empresa no cuca
            texto = '{};{};{};{};{};{};Empresa não encontrada no DPCUCA'.format(cod, cnpj, nome, nome_socio, data, valor)
            if not _verificar_empresa(cnpj, andamentos, texto, 'dpcuca'):
                continue

        # Entra na opção de cálculos do DPCUCA
        evento, imagen = calculos()

        # verifica se tem funcionário cadastrado
        if _find_img('Sem{}.png'.format(imagen), conf=0.95):
            _click_img('Sair.png', conf=0.9)
            _inicial('DPCUCA')
            _escreve_relatorio_csv('{};{};{};{};{};{};Sem {} cadastrado'.format(cod, cnpj, nome, nome_socio, data, valor, evento), nome=andamentos)
            print('❌ Sem {} cadastrado'.format(evento))
            continue

        # Adiciona o salário
        adicionar(andamentos, empresa)
        time.sleep(1)
        p.press('esc', presses=3)


@_time_execution
def run():
    usuario = p.prompt(text='Qual o usuario do robô?', title='Script incrível')
    empresas = _open_lista_dados()
    andamentos = 'Insere Salário Pro-Labore'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('DPCUCA')
    if entrar_na_empresa(index, empresas, andamentos, usuario) == 'OK':
        print('Informações adicionadas.')
    _fechar()


if __name__ == '__main__':
    run()
