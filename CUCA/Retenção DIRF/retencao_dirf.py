# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def relatorio_retencoes(empresa):
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(r'{}\{}'.format(e_dir, 'Relatórios'), exist_ok=True)

    cod, cnpj, nome = empresa
    p.press(['tab', '2'], interval=0.5)
    time.sleep(0.5)
    p.hotkey('ctrl', 'p')
    time.sleep(2)
    controle = ''
    # Espera o PDF abrir
    while not _find_img('PDF.png', conf=0.5):
        time.sleep(1)
        if _find_img('NaoTemRetencao.png', conf=0.9) or _find_img('NaoTemRetencao2.png', conf=0.9):
            p.press('enter')
            controle = 'n'
            break
    if controle == 'n':
        return False
    # Salvar o PDF
    _click_img('PDF.png', conf=0.9)
    _wait_img('SalvarPDF.png', conf=0.9, timeout=-1)
    time.sleep(1)
    # Usa o pyperclip porque o pyautogui não digita letra com acento
    texto = ' - '.join([nome, cod, 'Relatório Conta Corrente Retenções.pdf'])
    pyperclip.copy(texto)
    p.hotkey('ctrl', 'v')
    time.sleep(1)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\CUCA\Retenção DIRF\{}\{}'.format(e_dir, 'Relatórios'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')

    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('SalvarComo.png', 0.9):
        p.press('s')
        time.sleep(1)
    while _find_img('Progresso.png', 0.9):
        time.sleep(0.5)
    while _find_img('PDF.png', 0.9):
        p.press('esc')
        time.sleep(1)
        if _find_img('Comunicado.png', 0.9):
            p.press(['right', 'enter'], interval=0.2)
            time.sleep(1)
    return True


def gerar_dirf():
    p.hotkey('ctrl', 'g')
    time.sleep(0.5)
    if _find_img('NaoTemDIRF.png', conf=0.9):
        p.press('enter')
        p.press('esc', presses=2, interval=0.5)
        return False
    _wait_img('GerarDIRF.png', conf=0.9, timeout=-1)
    p.press(['n', 'tab', 'n', 'tab', 'tab'], interval=0.5)
    time.sleep(0.5)
    p.write('T')
    time.sleep(0.5)
    p.press('enter')
    _wait_img('ExportacaoCompleta.png', conf=0.9, timeout=-1)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    p.press('esc', presses=2, interval=0.5)
    return True


def consulta_retencao(index, empresas, andamentos):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        cod, cnpj, nome = empresa

        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('dpcuca')
            p.getWindowsWithTitle('DPCUCA')[0].maximize()

        _inicial('CUCA')

        # Verificações de login e empresa
        if not _login(empresa, 'Codigo', 'cuca', andamentos, '8', datetime.now().year):
            continue
        texto = '{};{};{};Robô sem acesso no CUCA'.format(cod, cnpj, nome)
        if not _verificar_empresa(cnpj, andamentos, texto, 'cuca'):
            continue

        # Abrir tela de C/C Retenções
        p.hotkey('alt', 'l')
        p.press(['a', 'c', 'c', 'enter'], interval=0.5)
        _wait_img('CCRetencoes.png', conf=0.9, timeout=-1)

        if not relatorio_retencoes(empresa):
            print('** Não existe registro para imprimir **')
            r = 'Não existe registro para imprimir'
        else:
            print('>>> Registro de retenção salvo <<<')
            r = 'Registro de retenção salvo'

        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, r]), nome=andamentos)

        if not gerar_dirf():
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe registro para exportar!', r]),
                                   nome=andamentos)
            print('** Não existe registro para exportar! **')
        else:
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Exportação completada com sucesso!', r]),
                                   nome=andamentos)
            print('>>> Exportação completada com sucesso! <<<')

        _inicial('CUCA')


@_time_execution
def run():
    empresas = _open_lista_dados()
    andamentos = 'Retenção DIRF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('CUCA')
    consulta_retencao(index, empresas, andamentos)
    _fechar()


if __name__ == '__main__':
    run()
