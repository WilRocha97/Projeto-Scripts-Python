# -*- coding: utf-8 -*-
from dados import empresas
from datetime import datetime
import pyautogui as p
import time

import sys
sys.path.append('..')
from comum import _horario, _iniciar_atalho, _escolher, _fechar, _login, _escrever, _verificar_empresa, _inicial, _esperar, _clicar

contas = [

]


def implantacao():
    for count, empresa in enumerate(empresas, start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'CUCA'):
            _iniciar_atalho()
            _escolher('cuca')

        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
        cod, cnpj, nome, valor, comp, ano = empresa

        print(' - '.join([indice, cod, cnpj, nome]))
        # Defini os textos usados nas respostas das verificações de login e empresa
        if not _login('12', datetime.now().year, empresa, 'Codigo', 'cuca'):
            p.press('enter')
            _escrever('Crédito PER DCOMP', ';'.join([cod, cnpj, nome, 'Empresa Inativa']))
            print('** Sem sócio cadastrado **\n')
            continue
        if not _verificar_empresa(cnpj, 'cuca'):
            _escrever('Crédito PER DCOMP', ';'.join([cod, cnpj, nome, 'Robô sem acesso no CUCA']))
            print('** Robô sem acesso no CUCA **\n')
            continue

        # Abrir Razão
        p.hotkey('alt', 'l')
        time.sleep(0.2)
        p.press('c')
        time.sleep(0.2)
        p.press('r')
        time.sleep(0.2)
        p.press('enter')
        _esperar(r'imagens\LivroRazao.png')
        _clicar(r'imagens\Conferencia.png')
        time.sleep(1)
        p.hotkey('ctrl', 'p')
        _esperar(r'imagens\Contas.png')
        if p.locateOnScreen(r'imagens\NaoContas.png'):
            _escrever('Crédito PER DCOMP', ';'.join([cod, cnpj, nome, 'Sem conta cadastrada']))
            print('** Sem conta cadastrada **\n')
            _inicial('cuca')
            continue

        # Selecionar as contas para gerar o relatório
        for conta in contas:
            _clicar(r'imagens\Numero.png')
            time.sleep(0.5)
            _clicar(r'imagens\NumeroDaConta.png')
            p.press('down')
            p.write(conta)
            time.sleep(1)
            p.hotkey('ctrl', 'a')
            time.sleep(1)

        p.press('enter')

        _inicial('cuca')


if __name__ == '__main__':
    try:
        inicio = datetime.now()

        # Verifica se o CUCA já está aberto, se não ele abre
        p.hotkey('win', 'm')
        if not p.locateOnScreen(r'imagens\CUCA.png', confidence=0.9):
            _iniciar_atalho()
            _escolher('cuca')
        else:
            p.getWindowsWithTitle('(SP)')[0].maximize()
            time.sleep(2)

        implantacao()
        _fechar()
        print(datetime.now() - inicio)
    except SystemExit:
        p.alert(title='Script incrível', text='Script encerrado.')
    except p.FailSafeException:
        p.alert(title='Script incrível', text='Encerrado manualmente.')
    except():
        p.alert(title='Script incrível', text='ERRO')
