# -*- coding: utf-8 -*-
from dados import empresas
from datetime import datetime
import pyautogui as p
import time
import sys

sys.path.append('..')
from comum import _horario, _fechar, _escrever, _verificar_empresa, _inicial, _esperar, _clicar, _iniciar, _login


def reprocesso():
    p.hotkey('alt', 'd')
    _esperar(r'imagens\SelecionarReprocessar.png')
    p.press('r')
    _esperar(r'imagens\DesejaReprocessar.png')
    p.hotkey('alt', 's')
    _esperar(r'imagens\PreencherMotivo.png')
    p.write('A')
    time.sleep(0.5)
    p.hotkey('alt', 'c')
    while not p.locateOnScreen(r'imagens\ReprocessoConcluido.png'):
        time.sleep(1)
        p.moveTo(683, 384)
        p.moveTo(684, 383)
    _clicar(r'imagens\ReprocessoConcluido.png')
    p.press('enter')
    time.sleep(1)
    _escrever('Tipo de Apuração', 'Reprocesso concluído')
    print('>>> Reprocesso concluído <<<\n')
    if p.locateOnScreen(r'imagens\Criticas.png'):
        _clicar(r'imagens\Criticas.png')
        p.press('esc')


def alterar_tipo():
    for count, empresa in enumerate(empresas, start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # verificar horário
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
            _escrever('Tipo de Apuração', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa']))
            print('** Não é possível entrar na empresa **\n')
            continue
        if not _verificar_empresa(cnpj, 'cuca'):
            _escrever('Tipo de Apuração', ';'.join([cod, cnpj, nome, 'Empresa não encontrada no CUCA']))
            print('** Empresa não encontrada no CUCA **\n')
            continue

        # Abrir empresas
        while not p.locateCenterOnScreen(r'imagens\CadastroDeEmpresas.png'):
            p.hotkey('ctrl', 'e')
            time.sleep(5)
            if p.locateOnScreen(r'imagens\Logo.png'):
                p.press('esc')
        time.sleep(2)
        if p.locateOnScreen(r'imagens\Logo.png'):
            p.press('esc')
        # Datas/Livros
        p.hotkey('ctrl', 'd')
        _esperar(r'imagens\DatasTermosLivrosOutros.png')

        # Conta Corrente Retenções
        _clicar(r'imagens\ContaCorrenteRetencoes.png')
        _esperar(r'imagens\MesDeInicio.png')
        if p.locateCenterOnScreen(r'imagens\RegimeDeCompetencia.png'):
            _escrever('Tipo de Apuração', ';'.join([cod, cnpj, nome, 'Não precisa alterar']))
            print('** Não precisa alterar **\n')
            _inicial('cuca')
            continue

        p.press('enter')
        time.sleep(1)
        if p.locateOnScreen(r'imagens\BalancoFechado.png'):
            p.press('enter')
            _escrever('Tipo de Apuração', ';'.join([cod, cnpj, nome, 'Balanço fechado']))
            print('** Balanço fechado **\n')
            _inicial('cuca')
            continue

        p.doubleClick(p.locateCenterOnScreen(r'imagens\TipoIRRF.png'))
        time.sleep(0.5)
        p.write('2')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write('2')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write('1')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.write('1')
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)
        if p.locateOnScreen(r'imagens\InformePerfil.png'):
            p.press('enter')
            time.sleep(0.5)
            p.press('esc')
            _escrever('Tipo de Apuração', ';'.join([cod, cnpj, nome, 'Não é possível alterar']))
            print('** Não é possível alterar **\n')
            _inicial('cuca')
            continue

        _inicial('cuca')
        # Anotar na planilha de andamentos
        with open('Tipo de Apuração.csv', 'a') as f:
            f.write(';'.join([cod, cnpj, nome, 'Alteração concluída;']))
        print('>>> alteração concluída <<<')

        _inicial('cuca')
        reprocesso()


if __name__ == '__main__':
    inicio = datetime.now()
    try:
        comp = datetime.now()
        mes = comp.strftime('%m')

        _iniciar('cuca')
        alterar_tipo()
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
