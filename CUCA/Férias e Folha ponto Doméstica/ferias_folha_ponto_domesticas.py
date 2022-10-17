# -*- coding: utf-8 -*-
from sys import path
path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start

from datetime import datetime
import os, time, pyperclip, pyautogui as p


def login(empresa, execucao, mes, ano):
    cod, cpf, nome = empresa
    if not _login(empresa, 'Codigo', 'dpcuca', 'Férias e Folha Ponto Domésticas', mes, ano):
        p.press('enter')
        _escreve_relatorio_csv('{};{};Empresa Inativa'.format(cod, nome), nome=execucao)
        print('❌ Empresa Inativa')
        return False
    if not _verificar_empresa(cpf, 'dpcuca'):
        _escreve_relatorio_csv(f'{cod};{nome};Empresa não encontrada no DPCUCA', nome=andamentos)
        print('❌ Empresa não encontrada no DPCUCA')
        return False
    return True


def competencia():
    mes_ferias = 0
    ano_ferias = 0
    # Perguntar mês das férias, deve ser entre 1 e 12, sempre no mês que o robô for executado
    while mes_ferias < 1 or mes_ferias > 12:
        comp = p.prompt(text='Qual a competência das Férias?', title='Script incrível', default='00/0000')
        comp = comp.split('/')
        mes_ferias = int(comp[0])
        ano_ferias = int(comp[1])

    # Define o mês do ponto
    mes_ponto = mes_ferias + 1
    ano_ponto = int(ano_ferias)
    
    # Define o ano e o mês do ponto caso a férias sejam no mês 12
    if mes_ferias == 12:
        mes_ponto = 1
        ano_ponto += 1

    return str(mes_ponto), str(mes_ferias), str(ano_ponto), str(ano_ferias)


def selecionar_gerar_salvar(cod, nome, execucao, f, pdf, gerou, situacao1, situacao2):
    os.makedirs(r'{}\{}'.format(e_dir, f), exist_ok=True)
    # Selecionar todos os funcionários e gerar os arquivos
    while not _find_img('InfoFuncionarios.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        
    p.press('n')
    _wait_img('EscolherFuncionarios.png', conf=0.9, timeout=-1)
    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 'a')
    if f == 'Férias':
        _wait_img('NumeroDeDias.png', conf=0.9, timeout=-1)
        p.write('120')
        time.sleep(0.5)
        p.press('enter')
    else:
        _wait_img('EscolhaPeriodo.png', conf=0.9, timeout=-1)
        p.press('enter')
        time.sleep(1)

    t = 0
    # Esperar gerar o PDF
    while not _find_img('PDF.png', conf=0.9):
        time.sleep(1)
        if t > 10:
            time.sleep(2)
            p.press('esc')
            time.sleep(1)
            p.press('esc')
            _escreve_relatorio_csv('{};{}'.format(gerou, situacao2), end='', nome=execucao)
            situacao = situacao2.replace(';', '').replace('\n', '')
            print('❌ {}'.format(situacao))
            break
        t += 1

    # Salvar o PDF
    if _find_img('PDF.png', conf=0.9):
        _click_img('PDF.png', conf=0.9)
        _wait_img('SalvarPDF.png', conf=0.9, timeout=-1)
        # Usa o pyperclip porque o pyautogui não digita letra com acento
        pyperclip.copy('{} - {} {}.PDF'.format(nome, cod, pdf))
        p.hotkey('ctrl', 'v')
        time.sleep(0.2)

        # Selecionar local
        p.press('tab', presses=6)
        time.sleep(0.2)
        p.press('enter')
        time.sleep(0.2)
        pyperclip.copy('V:\Setor Robô\Scripts Python\CUCA\Férias e Folha ponto Doméstica\{}\{}'.format(e_dir, f))
        p.hotkey('ctrl', 'v')
        time.sleep(0.2)
        p.press('enter')

        # Salva o arquivo
        time.sleep(0.2)
        p.hotkey('alt', 'l')
        time.sleep(1)
        if _find_img('SalvarComo.png', conf=0.9):
            p.press('s')
        p.press('esc')
        time.sleep(2)
        p.press('esc')
        time.sleep(1)
        p.press('esc')
        time.sleep(1)
        _escreve_relatorio_csv('{};{}'.format(gerou, situacao1), end='', nome=execucao)
        situacao = situacao1.replace(';', '').replace('\n', '')
        print('✔ {}'.format(situacao))


def ferias(execucao, cod, nome):
    # Cálculos > Evento Funcionários > Aba Férias > Imprimir > Férias Vencida ou a Vencer
    _wait_img('CUCAInicial.png', conf=0.9, timeout=-1)
    p.getWindowsWithTitle('DPCUCA PLUS |')[0].activate()
    p.hotkey('alt', 'c')
    time.sleep(1)
    p.press(['e', 'e', 'enter'], interval=0.5)
    _wait_img('EventoFuncionarios.png', conf=0.9, timeout=-1)
    time.sleep(2)
    if _find_img('SemFuncionarios.png', conf=0.9):
        _click_img('Sair.png', conf=0.9)
        _inicial('DPCUCA')
        _escreve_relatorio_csv('{};{};Sem funcionário cadastrado'.format(cod, nome), end='', nome=execucao)
        print('❌ Sem funcionário cadastrado')
        return True
    _click_img(r'Ferias.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 'p')
    time.sleep(1)

    _click_img('FeriasVV.png', conf=0.9)
    time.sleep(1)
    if not _find_img('Config.png', conf=0.9):
        p.press('tab')
        time.sleep(1)
        if not _find_img('LimiteDeConcessao.png', conf=0.9):
            while not _find_img('LimiteDeConcessao.png', conf=0.9):
                p.press('p')
        time.sleep(1)
        p.press('tab')
        time.sleep(1)
        if not _find_img('Nao.png', conf=0.9):
            while not _find_img('Nao.png', conf=0.9):
                p.press('n')
        time.sleep(1)
    p.hotkey('ctrl', 'p')
    gerou = ';'.join([cod, nome])
    selecionar_gerar_salvar(cod, nome, execucao, 'Férias', ' - Férias', gerou, 'Gerou Férias', 'Não gerou Férias')
    return True


def folha_ponto(execucao, cod, nome):
    # Cálculos > Evento Funcionários > Aba Cargo/Horário Trabalho > Imprimir > Folha Ponto
    _wait_img('CUCAInicial.png', conf=0.9, timeout=-1)
    _click_img('CUCAInicial.png', conf=0.9)
    p.hotkey('alt', 'c')
    time.sleep(1)
    p.press(['e', 'e', 'enter'], interval=0.5)
    _wait_img('EventoFuncionarios.png', conf=0.9, timeout=-1)
    time.sleep(2)
    if _find_img('SemFuncionarios.png', conf=0.9):
        _click_img('Sair.png', conf=0.9)
        _inicial('DPCUCA')
        _escreve_relatorio_csv(';Sem funcionário cadastrado', nome=execucao)
        print('❌ Sem funcionário cadastrado')
        return False
    _click_img('CargaHoraria.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 'p')
    time.sleep(1)
    p.press('f')
    time.sleep(0.5)
    selecionar_gerar_salvar(cod, nome, execucao, 'Ponto', ' - Folha Ponto', '', 'Gerou Folha Ponto\n', 'Não gerou Folha Ponto\n')
    return True


def consultar(emp, index, execucao):
    mes_ponto, mes_ferias, ano_ponto, ano_ferias = competencia()

    # Abrir o DPCUCA
    p.hotkey('win', 'm')
    _iniciar('DPCUCA')

    total_empresas = emp[index:]
    for count, empresa in enumerate(emp[index:], start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verificar horário
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('DPCUCA')
            p.getWindowsWithTitle('DPCUCA PLUS |')[0].maximize()

        _indice(count, total_empresas, empresa)

        cod, cpf, nome = empresa

        # Logar e consultar as férias
        if not login(empresa, execucao, mes_ferias, ano_ferias):
            continue
        if not ferias(execucao, cod, nome):
            continue

        # Logar e consultar as folhas ponto
        if not login(empresa, execucao, mes_ponto, ano_ponto):
            continue
        if not folha_ponto(execucao, cod, nome):
            continue


@_time_execution
def run():
    execucao = 'Férias e Folha Ponto Domésticas'
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    consultar(empresas, index, execucao)
    _fechar()


if __name__ == '__main__':
    run()
