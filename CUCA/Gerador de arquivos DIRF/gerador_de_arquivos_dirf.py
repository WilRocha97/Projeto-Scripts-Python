# -*- coding: utf-8 -*-
from dados import empresas
from datetime import datetime
import pyautogui as p
import time
import sys

sys.path.append('..')
from comum import _horario, _esperar, _clicar, _login, _fechar, _verificar_empresa, _escrever, _inicial, _iniciar


def gerar_dirf():
    for count, empresa in enumerate(empresas, start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('dpcuca')
            p.getWindowsWithTitle('DPCUCA')[0].maximize()

        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
        cod, cnpj, nome, mes, ano = empresa
        print(' - '.join([indice, cod, cnpj, nome, mes, ano]))

        # Verificações de login e empresa
        if not _login(mes, ano, empresa, 'Codigo', 'dpcuca'):
            p.press('enter')
            _escrever('Gerador de arquivos DIRF', ';'.join([cod, cnpj, nome, mes, ano, 'Empresa Inativa']))
            print('** Empresa Inativa **\n\n')
            continue
        # CNPJ cpm os separadores para poder verificar a empresa no cuca
        if not _verificar_empresa(cnpj, 'dpcuca'):
            _escrever('Gerador de arquivos DIRF', ';'.join([cod, cnpj, nome, mes, ano, 'Robô sem acesso no DPCUCA']))
            print('** Robô sem acesso no DPCUCA **\n\n')
            continue

        time.sleep(4)
        _clicar(r'imgComum\TCUCAInicial.png')
        # Abrir Calculos
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press(['d', 'd', 'enter'], interval=0.5)
        _esperar(r'imagens\GerarInfo.png')
        _clicar(r'imagens\Pasta.png')
        _esperar(r'imagens\procurarPasta.png')
        time.sleep(0.5)
        p.press(['tab', 'tab', 'tab'], interval=0.5)
        time.sleep(0.5)
        p.write('T:\DIRF 2021-2020')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        if p.locateOnScreen(r'imagens\GerarBenefciarios.png'):
            p.click(r'imagens\GerarBenefciarios.png')
        time.sleep(0.5)
        p.click(r'imagens\Gerar.png')

        p.hotkey('alt', 'a')
        time.sleep(1)
        if p.locateOnScreen(r'imagens\Atencao.png'):
            p.press('enter')
        time.sleep(2)
        while not p.locateOnScreen(r'imagens\Inicio.png'):
            time.sleep(1)
            p.moveTo(683, 384)
            p.moveTo(684, 383)
            if p.locateOnScreen(r'imagens\OK.png'):
                p.press('enter')
                time.sleep(0.5)
            if p.locateOnScreen(r'imagens\Situacoes.png'):
                p.press('enter')
                time.sleep(0.5)
            if p.locateOnScreen(r'imagens\DIRFgerada.png'):
                p.press('enter')
                _escrever('Gerador de arquivos DIRF', ';'.join([cod, cnpj, nome, mes, ano, 'Arquivo DIRF gerado!']))
                print('>>> Arquivo DIRF gerado! <<<\n\n')
            if p.locateOnScreen(r'imagens\NaoDIRF.png'):
                p.press('enter')
                _escrever('Gerador de arquivos DIRF', ';'.join([cod, cnpj, nome, mes, ano, 'Não foram encontrados informações para DIRF!']))
                print('** Não foram encontrados informações para DIRF! **\n\n')

        _clicar(r'imagens\JanelaAGUARDE.png')
        p.press('esc')

        _inicial('dpcuca')


if __name__ == '__main__':
    # Perguntar qual relatório
    inicio = datetime.now()
    try:
        _iniciar("dpcuca")
        gerar_dirf()
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
