# -*- coding: utf-8 -*-
from dados import empresas
from datetime import datetime
import pyautogui as p
import time
import sys

sys.path.append('..')
from comum import _horario, _esperar, _login, _fechar, _verificar_empresa, _escrever, _inicial, _iniciar


def editar_cadastro():
    for count, empresa in enumerate(empresas, start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'CUCA'):
            _iniciar('cuca')
            p.getWindowsWithTitle('(SP)')[0].maximize()

        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
        cod, cnpj, nome = empresa

        print(' - '.join([indice, cod, cnpj, nome]))
        # Defini os textos usados nas respostas das verificações de login e empresa
        if not _login(mes, datetime.now().year, empresa, 'Codigo', 'cuca'):
            p.press('enter')
            _escrever('Cadastro Reinf', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa']))
            print('** Não é possível entrar na empresa **\n')
            continue
        if not _verificar_empresa(cnpj, 'cuca'):
            _escrever('Cadastro Reinf', ';'.join([cod, cnpj, nome, 'Robô sem acesso no CUCA']))
            print('** Robô sem acesso no CUCA **\n')
            continue

        p.hotkey('alt', 'l')
        time.sleep(0.5)
        p.press('s', presses=3)
        time.sleep(0.5)
        p.press('enter', presses=2)
        time.sleep(2)

        if p.locateOnScreen(r'imagens\ok.png'):
            _escrever('Cadastro Reinf', ';'.join([cod, cnpj, nome, 'Empresa editada']))
            print('>>> Empresa editada <<<\n')
            _inicial('CUCA')
            continue

        if p.locateOnScreen(r'imagens\SemClass.png'):
            _escrever('Cadastro Reinf', ';'.join([cod, cnpj, nome, 'Sem Classificação Tributária']))
            print('** Sem Classificação Tributária **\n')
            _inicial('CUCA')
            continue

        p.press('enter')
        time.sleep(0.5)
        p.write('5')
        time.sleep(0.5)
        while not p.locateOnScreen(r'imagens\Inicio.png'):
            p.press('tab')
            time.sleep(0.5)
        p.write('052021')
        time.sleep(0.5)
        p.press('tab', presses=2)
        time.sleep(0.5)
        while not p.locateOnScreen(r'imagens\Producao2.png'):
            p.write('1')
            time.sleep(1)
        p.press('enter')
        time.sleep(0.5)

        if p.locateOnScreen(r'imagens\Balanco.png'):
            _escrever('Cadastro Reinf', ';'.join([cod, cnpj, nome, 'Balanço fechado no mês']))
            print('** Balanço fechado no mês **\n')
            p.press('enter')
            _inicial('CUCA')
            continue

        p.press('s')
        _esperar(r'imagens\Referencia.png')

        Dados = ';'.join([cod, cnpj, nome, 'Empresa editada'])

        try:
            with open('Cadastro Reinf.csv', 'a') as f:
                f.write(Dados + '\n')
        except:
            with open('Cadastro Reinf - Auxiliar.csv', 'a') as f:
                f.write(Dados + '\n')

        print('>>> Empresa editada <<<\n')

        _inicial('CUCA')


if __name__ == '__main__':
    # mouse = p.mouseInfo()
    inicio = datetime.now()
    try:
        mes = '0'
        while mes == '0':
            mes = p.prompt(title='Script incrível', text='Qual o mês?', default='0')

        _iniciar('cuca')
        editar_cadastro()
        _fechar()

        print(datetime.now() - inicio)
    except SystemExit:
        print(datetime.now() - inicio)
        p.alert(title='Script incrível', text='Script encerrado.')
    except p.FailSafeException:
        print(datetime.now() - inicio)
        p.alert(title='Script incrível', text='Encerrado manualmente.')
    except():
        print(datetime.now() - inicio)
        p.alert(title='Script incrível', text='ERRO')
