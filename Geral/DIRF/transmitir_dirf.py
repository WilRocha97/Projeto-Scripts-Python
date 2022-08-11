# -*- coding: utf-8 -*-
import pyautogui as p
import time
from sys import path

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice


def esperar(imagen):
    while not p.locateOnScreen(imagen, confidence=0.9):
        time.sleep(1)


def clicar(imagen, tempo=0.1):
    try:
        p.click(p.locateCenterOnScreen(imagen, confidence=0.9))
        time.sleep(tempo)
        return True
    except:
        return False


def transmitir(empresa):
    cod, cnpj, nome, regime = empresa

    p.getWindowsWithTitle('Dirf 2022')[0].maximize()
    esperar(r'imagens2\RestaurarCopia.png')
    clicar(r'imagens2\RestaurarCopia.png')
    esperar(r'imagens2\Pesquisar.png')
    clicar(r'imagens2\Pesquisar.png')
    esperar(r'imagens2\Abrir.png')
    p.write('T:\DCTF JAQUE\DIRF 2022-2021\Copias\DIRF2022' + cnpj + '2021.DBK')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'a')
    time.sleep(1)

    if p.locateOnScreen(r'imagens2\ArquivoInvalido.png'):
        p.press('enter')
        time.sleep(0.5)
        p.hotkey('alt', 'c')
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Arquivo não encontrado']))
        print('>>> Arquivo não encontrado <<<')
        return False

    esperar(r'imagens2\Marcar.png')
    clicar(r'imagens2\Marcar.png')
    esperar(r'imagens2\Avancar.png')
    p.hotkey('alt', 'a')
    esperar(r'imagens2\Concluir.png')
    p.hotkey('alt', 'n')
    esperar(r'imagens2\RestaurarCopia.png')
    p.hotkey('ctrl', 'g')
    esperar(r'imagens2\AbrirDeclaracao.png')
    p.hotkey('alt', 'o')
    while not p.locateOnScreen(r'imagens2\Avancar.png'):
        if p.locateOnScreen(r'imagens2\ErroTransmitir.png', confidence=0.9):
            p.hotkey('alt', 'c')
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Erro']))
            print('>>> Erro <<<')

            # Deletar dados
            p.hotkey('alt', 'd')
            esperar(r'imagens2\Excluir.png')
            clicar(r'imagens2\Excluir.png')
            esperar(r'imagens2\ExcluirDeclara.png')
            clicar(r'imagens2\Marcar.png')
            esperar(r'imagens2\OK.png')
            p.hotkey('alt', 'o')
            esperar(r'imagens2\DesejaExcluir.png')
            p.hotkey('alt', 's')
            return False
        time.sleep(0.5)
    p.hotkey('alt', 'a')
    esperar(r'imagens2\Local.png')
    time.sleep(1)
    p.write('sp')
    time.sleep(0.5)
    esperar(r'imagens2\Avancar.png')
    p.hotkey('alt', 'a')

    time.sleep(1)

    p.press('s')
    esperar(r'imagens2\Transmitir.png')
    clicar(r'imagens2\Sim.png')
    if regime == 'SN':
        if p.locateOnScreen(r'imagens2\TemCertificado.png'):
            clicar(r'imagens2\TemCertificado.png')
            time.sleep(0.5)
            esperar(r'imagens2\Avancar.png')
            p.hotkey('alt', 'a')

    esperar(r'imagens2\Resultado.png')
    time.sleep(2)
    if p.locateOnScreen(r'imagens2\ErroTransmitir.png'):
        p.hotkey('alt', 'n')
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Erro']))
        print('>>> Erro <<<')

        # Deletar dados
        p.hotkey('alt', 'd')
        esperar(r'imagens2\Excluir.png')
        clicar(r'imagens2\Excluir.png')
        esperar(r'imagens2\ExcluirDeclara.png')
        clicar(r'imagens2\Marcar.png')
        esperar(r'imagens2\OK.png')
        p.hotkey('alt', 'o')
        esperar(r'imagens2\DesejaExcluir.png')
        p.hotkey('alt', 's')
        return False

    p.hotkey('alt', 'n')
    esperar(r'imagens2\Imprimir.png')
    clicar(r'imagens2\SelecionaImpr.png')
    time.sleep(0.5)
    p.press('down', presses=2)
    time.sleep(0.5)
    p.press('enter')
    clicar(r'imagens2\BotaoImprimir.png')
    esperar(r'imagens2\SalvarRecibo.png')

    # Salvar o PDF
    time.sleep(0.5)
    p.write(nome)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    if p.locateCenterOnScreen(r'imagens2\SalvarComo.png', confidence=0.9):
        p.press('s')
        time.sleep(0.2)
    esperar(r'imagens2\RestaurarCopia.png')
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Recibo salvo']))
    print('>>> Recibo Salvo <<<')

    # Deletar dados
    p.hotkey('alt', 'd')
    esperar(r'imagens2\Excluir.png')
    clicar(r'imagens2\Excluir.png')
    esperar(r'imagens2\ExcluirDeclara.png')
    clicar(r'imagens2\Marcar.png')
    esperar(r'imagens2\OK.png')
    p.hotkey('alt', 'o')
    esperar(r'imagens2\DesejaExcluir.png')
    p.hotkey('alt', 's')
    return True


@_time_execution
def run():
    p.hotkey('win', 'm')
    empresas = _open_lista_dados()
            
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    time.sleep(0.5)
    p.getWindowsWithTitle('Dirf')[0].maximize()

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        _indice(count, total_empresas, empresa)

        if not transmitir(empresa, ):
            continue


if __name__ == '__main__':
    run()
