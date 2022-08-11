# -*- coding: utf-8 -*-
from dados import empresas
from datetime import datetime
import pyautogui as p
import time
import sys

sys.path.append('..')
from comum import _horario, _fechar, _login, _escrever, _verificar_empresa, _inicial, _iniciar, _esperar, _clicar


def login(empresa):
    cod, cnpj, nome = empresa
    # Defini os textos usados nas respostas das verificações de login e empresa
    if not _login(12, datetime.now().year, 12, datetime.now().year, empresa, 'CNPJ', 'cuca'):  # recebe o ano do dado anterior caso seja igual não perde tempo trocando no cuca
        time.sleep(1)
        if p.locateOnScreen(r'imagens\SemSocio.png', confidence=0.9):
            p.press('enter')
            _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa não existe sócio cadastrado']))
            print('** Não é possível entrar na empresa não existe sócio cadastrado **\n')
        if p.locateOnScreen(r'imagens\OpcaoInvalida.png', confidence=0.9):
            p.press('enter')
            _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa Opção inválida']))
            print('** Não é possível entrar na empresa Opção inválida **\n')
        if p.locateOnScreen(r'imagens\NaoLiberada.png', confidence=0.9):
            p.press('enter')
            _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa não liberada no período selecionado!']))
            print('** Não é possível entrar na empresa não liberada no período selecionado! **\n')
        if p.locateOnScreen(r'imgComum\EscolherEmpresaCUCA.png'):
            p.click(p.locateCenterOnScreen(r'imgComum\EscolherEmpresaCUCA.png'))
            time.sleep(1)
            p.press('esc')
            _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Acesso bloqueado']))
            print('** Acesso bloqueado **\n')
        else:
            p.press('enter')
            _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Não é possível entrar na empresa']))
            print('** Não é possível entrar na empresa **\n')
        return False
    if not _verificar_empresa(cnpj, 'cuca'):
        _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Empresa não encontrada no CUCA']))
        print('** Empresa não encontrada no CUCA **\n')
        return False
    return True


def  editar_dados():
    for count, empresa in enumerate(empresas, start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'CUCA'):
            _iniciar('cuca')
            p.getWindowsWithTitle('(SP)')[0].maximize()

        # Cria um indice par saber qual linha dos dados está
        indice = '[ ' + str(count) + ' de ' + str(len(empresas)) + ' ]'
        cod, cnpj, nome = empresa
        _inicial('cuca')

        print(' - '.join([indice, cod, cnpj, nome]))

        if not login(empresa):
            continue

        # Abrir tela de exportação e importação
        while not p.locateOnScreen(r'imagens\ref.png', confidence=0.9):
            p.hotkey('alt')
            time.sleep(2)
        while not p.locateOnScreen(r'imagens\export.png', confidence=0.9):
            p.press('right')

        # exportar informações do esped contabil fiscal(ecf)
        p.press('enter', presses=2)
        time.sleep(1)
        p.press('e', presses=7)
        p.press('enter')
        time.sleep(1)
        _esperar(r'imagens\Dados.png')
        time.sleep(2)

        # clicar em dados cadastrais
        _clicar(r'imagens\Dados.png')
        time.sleep(1)

        # editar os dados
        if not p.locateOnScreen(r'imagens\Criterio.png') or not p.locateOnScreen(r'imagens\TipoEscr.png'):
            p.press('enter')
            _esperar(r'imagens\habilitado.png')
            p.press('tab')
            time.sleep(0.5)
            while not p.locateOnScreen(r'imagens\LivroCaixa.png'):
                p.press('l')
                time.sleep(0.5)
            p.press('tab')
            time.sleep(0.5)
            while not p.locateOnScreen(r'imagens\RegimeComp.png'):
                p.press('2')
                time.sleep(0.5)
            time.sleep(0.5)
            p.press('enter')
            time.sleep(2)

        # clicar em geração sped ecf
        _clicar(r'imagens\GeraECF.png')
        time.sleep(2)

        # gerar arquivo
        p.press('enter')
        time.sleep(1)
        if p.locateOnScreen(r'imagens\DesejaContinuar.png'):
            p.press('s')
        while not p.locateOnScreen(r'imagens\ArquivoGerado.png'):
            if p.locateOnScreen(r'imagens\Advertencias.png'):
                p.press('esc')
            time.sleep(2)
        p.press('enter')
        _esperar(r'imagens\Arquivo.png')
        p.press('esc')

        _escrever('Dados Cadastrais SPED ECF', ';'.join([cod, cnpj, nome, 'Arquivo gerado']))

        print('>>> Consulta concluída <<<\n')
        _inicial('CUCA')


if __name__ == '__main__':
    inicio = datetime.now()
    try:
        comp = datetime.now()
        mes = comp.strftime('%m')

        _iniciar('cuca')
        editar_dados()
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
