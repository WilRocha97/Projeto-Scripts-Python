# -*- coding: utf-8 -*-
from datetime import datetime
from dateutil.relativedelta import relativedelta
from time import sleep
from pyperclip import copy
from os import makedirs
from tkinter import *
from ttk import Combobox
from pyautogui import press, hotkey, write, click, alert, getWindowsWithTitle

from sys import path
path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def escolher_relatorio():
    root = Tk()
    root.title('Script incrível')

    window_width = 400
    window_height = 100
    # find the center point
    center_x = int(root.winfo_screenwidth() / 2 - window_width / 2)
    center_y = int(root.winfo_screenheight() / 2 - window_height / 2)

    # set the position of the window to the center of the screen
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

    relatorio = StringVar(root)
    relatorio.set("Qual relatório?")  # default value

    relatorios = ["Férias vencidas ou a vencer", "Ficha de Registro de Funcionários", "Folha de Pagamento - Impressão Multipla", "Holerites - Adiantamento",
                  "Holerites - Mensal", "Holerite Pro Labore", "Provisões 13º e Ferias", "Relatório de empresas", "Relatório de Experiência geral", "Relatório Geral - Impressão Multipla",
                  "Rescisão", "Resumo Geral Mensal - Impressão Multipla", "Resumo Mensal - Impressão Multipla", "Vale Transporte"]

    # drop down menu
    Combobox(root, textvariable=relatorio, values=relatorios, width=55).pack()

    # botão com comando para fechar o tkinter
    Button(root, text="Ok", command=root.destroy).pack()

    root.mainloop()

    return str(relatorio.get())


def calculos(relatorio):
    # Entrar na opção de Cálculos
    sleep(2)
    _click_img('CUCAInicial.png', conf=0.9)
    hotkey('alt', 'c')
    sleep(0.5)
    while not _find_img('SelecionarEvento.png', conf=0.9):
        _click_img('CUCAInicial.png', conf=0.9)
        hotkey('alt', 'c')
        sleep(0.5)

    # Algumas consultas vão aqui: Evento Sócios
    if relatorio == 'Informe de Rendimento - Sócios' or relatorio == 'Holerite Pro Labore':
        evento, imagen, cont = 'sócio', 'Socios', 3

    # O restante vai em eventos funcionários
    else:
        evento, imagen, cont = 'funcionário', 'Funcionarios', 2

    press('e', presses=cont, interval=0.2)
    press('enter')
    _wait_img('Eventos{}.png'.format(imagen), conf=0.9, timeout=-1)
    sleep(3)
    return evento, imagen, cont


def gerar(relatorio, andamentos, empresa):
    # comando para gerar arquivos no dpcuca e faz uma verificação se pode gerar
    if _find_img('NaoImprimi.png', conf=0.9):
        cod, cnpj, razao, mes, ano = empresa
        _inicial('DPCUCA')
        sleep(0.5)
        _escreve_relatorio_csv(f'{cod};{cnpj};{razao};{mes};{ano};Não é possível gerar {relatorio}', nome=andamentos)
        print('❌ Não é possível gerar {}'.format(relatorio))
        return False
    hotkey('alt', 'p')
    return True


def imprimir(relatorio, andamentos, empresa, texto, espera=10, diretorio='Relatórios'):
    cod, cnpj, razao, mes, ano = empresa
    # cria uma pasta com o nome do relatório para salvar os arquivos
    if relatorio != 'Provisões 13º e Ferias':
        makedirs(r'{}\{}'.format(e_dir, diretorio), exist_ok=True)

    # Esperar o PDF gerar
    sleep(1)
    x = 0
    while not _find_img('PDF.png', conf=0.9):
        sleep(1)
        x += 1
        # exclusivo para rescisão
        if _find_img('Homologonet.png'):
            press('enter')
        if x > espera:
            if not _find_img('PDF.png', conf=0.9):
                # Se for 'Provisões' não precisa entrar aqui, essa execução já tem verificações próprias
                if relatorio != 'Provisões':
                    nao_gerou(relatorio, andamentos, empresa)
                return False

    # Salvar o PDF
    while not _find_img('SalvarPDF.png', conf=0.9):
        if relatorio == 'Empresas':
            _click_img('PDF2.png', conf=0.9)
            sleep(5)
            break
        else:
            click(109, 68)
    sleep(1)

    # Usa o pyperclip porque o pyautogui não digita letra com acento
    copy(texto)
    hotkey('ctrl', 'v')
    sleep(0.5)

    # Selecionar local
    press('tab', presses=6)
    sleep(0.5)
    press('enter')
    sleep(0.5)

    if relatorio == 'Relatório de Experiência geral':
        copy('V:\Setor Robô\Scripts Python\Geral\Separa Experiência a Vencer\PDF')

    elif relatorio == 'Relatório de empresas':
        copy('V:\Setor Robô\Relatorios\pdf_dp')

    # Coloca os PDFs das Provisões direto na pasta do robô
    elif relatorio == 'Provisões 13º e Ferias':
        if int(mes) < 10:
            mes = f'0{str(mes)}'
        diretorio = r'\\vpsrv03\Arq_Robo\Provisões 13º e Férias\{}\Provisões {}-{}\Provisões'.format(ano, str(mes), ano)
        makedirs(diretorio, exist_ok=True)
        copy(diretorio)

    # Caminho padrão dos relatórios
    else:
        copy('V:\Setor Robô\Scripts Python\CUCA\Relatórios DP-CUCA\{}\{}'.format(e_dir, diretorio))

    # cola o caminho para salvar
    hotkey('ctrl', 'v')
    sleep(0.5)
    press('enter')

    # salva o arquivo na pasta
    sleep(0.5)
    hotkey('alt', 'l')
    sleep(1)
    if _find_img('SalvarComo.png', 0.9):
        press('s')
        sleep(1)

    # espera terminar de salvar o arquivo
    while _find_img('Progresso.png', 0.9):
        sleep(0.5)

    # fecha a visualização
    while _find_img('PDF.png', 0.9):
        press('esc')
        sleep(1)
        if _find_img('Comunicado.png', 0.9):
            press(['right', 'enter'], interval=0.2)
            sleep(1)

    # se for alguma dessas consultas não escreve na planilha nesse momento
    for i in ['Provisões 13º e Ferias', 'Relatório de empresas', 'Relatório de Experiência geral']:
        if relatorio == i:
            return True

    print('✔ {} gerado'.format(relatorio))
    _escreve_relatorio_csv(f'{cod};{cnpj};{razao};{mes};{ano};{relatorio} gerado', nome=andamentos)
    return True


def holerite(relatorio, andamentos, empresa):
    # Gerar holerite
    if not gerar(relatorio, andamentos, empresa):
        return False
    while not _find_img('SelecionaHoleriteA.png', 0.9):
        sleep(1)
        hotkey('alt', 'p')
        # Aviso que às vezes aparece
        if _find_img('RefaserCalculo.png', 0.9):
            _click_img('RefaserCalculo.png')
            press('enter')
            sleep(1)
    press(['h', 'enter'], interval=0.5)
    selecionar_funcionarios()
    _wait_img('MensagensHolerite.png', conf=0.9, timeout=-1)
    hotkey('alt', 'o')
    sleep(1)
    hotkey('alt', 'o')
    sleep(1)
    while _find_img('De.png', 0.9):
        if _find_img('RefaserCalculo.png', 0.9):
            _click_img('RefaserCalculo.png')
            press('enter')
    sleep(2)
    return True


def selecionar_funcionarios():
    # Selecionar todos os funcionários
    _wait_img('InfoFuncionarios.png', conf=0.9, timeout=-1)
    press('n')
    _wait_img('EscolherFuncionarios.png', conf=0.9, timeout=-1)
    hotkey('alt', 'o')
    sleep(0.7)
    hotkey('alt', 'a')


def nao_gerou(relatorio, andamentos, empresa, texto='Não gerou'):
    # a variável texto tem um padrão geral devido-as rescisões que usam um texto diferente
    cod, cnpj, nome, mes, ano = empresa
    _click_img('Sair.png', conf=0.9)
    _inicial('DPCUCA')
    _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Não gerou {relatorio}', nome=andamentos)
    print('❌ {} {}'.format(texto, relatorio))


def relatoriozinhos(relatorio, andamentos, empresa):
    # Os dois primeiros não precisa usar índice e nem a planilha de dados
    if relatorio == 'Relatório de Experiência geral':
        try:
            _login(empresa, 'Codigo', 'dpcuca')
        except:
            alert(title='Script incrível', text='Erro no login do DPCUCA')

        calculos(relatorio)
        # Aba Funcionários > Imprimir > Registro de Empregado
        _click_img('Funcionarios.png')
        _wait_img('Imprimir.png', conf=0.9, timeout=-1)
        sleep(1)
        gerar(relatorio, andamentos, empresa)
        _wait_img('Experiencia.png', conf=0.9, timeout=-1)
        press('e', presses=3)
        press('enter')
        press('p', presses=2)
        press('enter')

        _wait_img('InfoFuncionarios.png', conf=0.9, timeout=-1)
        press('n')
        _wait_img('EscolherFuncionarios.png', conf=0.9, timeout=-1)
        hotkey('alt', 'm')
        sleep(10)
        while _find_img('Carregando.png'):
            sleep(1.5)
        hotkey('alt', 'a')
        _wait_img('planilha.png', conf=0.9, timeout=-1)
        press('n')
        imprimir(relatorio, andamentos, empresa, 'Experiência.PDF', espera=1500, diretorio=relatorio)
        _inicial('DPCUCA')
        return True

    elif relatorio == 'Relatório de empresas':
        try:
            _login(empresa, 'Codigo', 'dpcuca')
        except:
            alert(title='Script incrível', text='Erro no login do DPCUCA')

        # formata a data para adicionar barra entre os números e com a quantidade certa de caracteres em cada casa
        data = '{:02d}/{:02d}/{:04d}'.format(datetime.now().day, datetime.now().month, datetime.now().year)

        hotkey('alt', 'e')
        _wait_img('Empresas.png', conf=0.9, timeout=-1)
        hotkey('alt', 'p')
        _wait_img('Relatorios.png', conf=0.9, timeout=-1)
        press('g')
        _wait_img('Gerador.png', conf=0.9, timeout=-1)
        _click_img('Exportacao.png' or 'Exportacao2.png', conf=0.9)
        sleep(1)
        hotkey('alt', 'p')
        _wait_img('EscolheEmpresas.png', conf=0.9, timeout=-1)
        hotkey('alt', 'm')
        _wait_img('listavazia.png', conf=0.9, timeout=-1)
        hotkey('alt', 'a')
        _wait_img('data.png', conf=0.9, timeout=-1)
        write(data)
        sleep(1)
        hotkey('alt', 'o')
        _wait_img('impressao.png', conf=0.9, timeout=-1)
        _click_img('impressao.png')
        press('s')
        if not imprimir(relatorio, andamentos, empresa, 'dpcuca - {:02d}{:02d}{:04d}.PDF'.format(datetime.now().day, datetime.now().month, datetime.now().year), espera=5000, diretorio=relatorio):
            return False
        _inicial('DPCUCA')
        return True

    cod, cnpj, nome, mes, ano = empresa

    # Texto padrão para digitar no nome do arquivo
    texto = ' - '.join([nome, cod, relatorio, mes, ano]) + '.PDF'

    if relatorio == 'Férias vencidas ou a vencer':
        # Aba Férias > Férias Vencidas ou a Vencer > Ordem de Limite de Concessão
        sleep(5)
        _click_img('Ferias.png', conf=0.9)
        sleep(2)
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('RelatorioDeFerias.png', conf=0.9, timeout=-1)
        if not _find_img('FeriasVencidasSelect.png'):
            _click_img('FeriasVencidas.png')
            sleep(1)
        hotkey('ctrl', 'p')
        selecionar_funcionarios()
        _wait_img('NumeroDeDias.png', conf=0.9, timeout=-1)
        write('120')
        sleep(0.5)
        press('enter')
        sleep(2)
        press('n')
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Ficha de Registro de Funcionários':
        # Aba Funcionários > Imprimir > Registro de Empregado
        _click_img('Funcionarios.png')
        _wait_img('Imprimir.png', conf=0.9, timeout=-1)
        sleep(1)
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('SelecionarRegistro.png', conf=0.9, timeout=-1)
        press(['r', 'enter'], interval=0.5)
        selecionar_funcionarios()
        _wait_img('OrdemDoRegistro.png', conf=0.9, timeout=-1)
        if _find_img('OrdemDoRegistro.png', 0.9):
            press('right')
            sleep(0.5)
            hotkey('ctrl', 'p')
            sleep(0.5)
            hotkey('alt', 'c')
        if not imprimir(relatorio, andamentos, empresa, texto, espera=120, diretorio=relatorio):
            return False

    elif relatorio == 'Holerites 13':  # Em análise
        # Aba secundária Adiantamento Obrigatório
        _click_img('AdiantamentoObrigatorio.png')
        if not holerite(relatorio, andamentos, empresa):
            return False
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Holerites - Adiantamento':
        # Aba secundária Adiantamento Obrigatório
        _click_img('AdiantamentoObrigatorio.png')
        if not holerite(relatorio, andamentos, empresa):
            return False
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Holerites - Mensal':
        if not holerite(relatorio, andamentos, empresa):
            return False
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Holerite Pro Labore':
        # Imprimir > Holerite
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('SelecionaHoleriteP.png', conf=0.9, timeout=-1)
        press(['h', 'enter'], interval=0.5)
        _wait_img('MensagensHolerite.png', conf=0.9, timeout=-1)
        hotkey('alt', 'o')
        # Clicar na opção 2 para gerar de todos os sócios e depois clicar para gera duas vias
        # Podem aparecer duas telas diferentes para configurar isso
        _wait_img('Opcao2.png', conf=0.99, timeout=-1)
        _click_img('Opcao2.png', conf=0.9)
        while not _find_img('DuasVias.png', conf=0.9):
            if _find_img('ComoImprimirHolerite.png', conf=0.9):
                hotkey('alt', 'o')
                sleep(2)
                break
        if _find_img('DuasVias.png', conf=0.9):
            _click_img('DuasVias.png', conf=0.9)
            _wait_img('EnviaRelatorio.png', conf=0.9, timeout=-1)
            hotkey('alt', 'a')
            sleep(2)
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Provisões 13º e Ferias':
        # Aba secundária Provisões e verifica se pode calcular provisões
        hotkey('alt', 'v')
        sleep(2)
        if not _find_img('Calcular.png', conf=0.9):
            _inicial('DPCUCA')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Não é possível calcular Provisões', nome=andamentos)
            print('❌ Não é possível calcular Provisões')
            return False

        # Calcular provisões
        hotkey('alt', 'c')
        while not _find_img('Confirmar.png', 0.9):
            sleep(5)
            if _find_img('JaExisteProvisao.png', conf=0.9):
                _click_img('JaExisteProvisao.png', conf=0.9)
                press('enter')
                sleep(1)
                hotkey('alt', 'd')
                _wait_img('DesfazTodos.png', conf=0.9, timeout=-1)
                _click_img('DesfazTodos.png', conf=0.9)
                sleep(1)
                press('s')
                sleep(1)
                press('s')
                sleep(1)
                press('t')
                sleep(5)
                hotkey('alt', 'c')
                sleep(5)
        if _find_img('Confirmar.png', 0.9):
            press('enter')
            _wait_img('CalcularParaTodos.png', conf=0.9, timeout=-1)
            _click_img('CalcularParaTodos.png', conf=0.9)
            _click_img('Todos.png', conf=0.9)

        # Imprimir > Clicar em Provisão 13.º > Imprimir, repetir o mesmo para provisão Férias
        provisoes = [("13", "Provisão 13º"), ("Ferias", "Provisão Férias"), ]
        controle = 0
        for provisao in provisoes:
            if not gerar(relatorio, andamentos, empresa):
                continue

            _wait_img('Provisao.png', conf=0.9, timeout=-1)
            while _find_img('Provisao{}.png'.format(provisao[0]), conf=0.9):
                _click_img('Provisao{}.png'.format(provisao[0]), conf=0.9)
                sleep(1)
                hotkey('ctrl', 'p')

            sleep(1)
            selecionar_funcionarios()
            sleep(1)
            while _find_img('Aguarde.png', conf=0.9):
                sleep(1)
            while _find_img('Processando.png', conf=0.9):
                sleep(1)
            # Textos específicos para essa execução
            if not imprimir(relatorio, andamentos, empresa, ' - '.join([nome, cod, provisao[1], mes, ano]) + '.PDF', espera=500):
                controle = 0
            else:
                if provisao[0] == '13':
                    controle += 1
                else:
                    controle += 2
        if controle == 0:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Não Gerou Provisão', nome=andamentos)
            print('❌ Não Gerou Provisões')
        elif controle == 1:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Gerou Provisão 13º', nome=andamentos)
            print('✔ Gerou Provisão 13º')
        elif controle == 2:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Gerou Provisão Férias', nome=andamentos)
            print('✔ Gerou Provisão Férias')
        elif controle == 3:
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Gerou Provisão 13º e Férias', nome=andamentos)
            print('✔ Gerou Provisão 13º e Férias')

        _inicial('DPCUCA')

    elif relatorio == 'Rescisão':
        # Aba secundária Rescisão Contrato > Imprimir > Rescisão
        hotkey('alt', 'r')
        sleep(2)
        # Verifica se tem funcionário na empresa
        if _find_img('SemFuncionarios.png', conf=0.9):
            nao_gerou(relatorio, andamentos, empresa, 'Não tem ')
            return False
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('SelecionarRescisao.png', conf=0.9, timeout=-1)
        press(['r', 'enter'], interval=0.5)
        _wait_img('SemFoto.png', conf=0.9, timeout=-1)
        _click_img('SemFoto.png')
        if _find_img('Rascunho.png'):
            _click_img('Rascunho.png')
        press('enter')
        if _find_img('AvisoRescisao.png'):
            press('enter')
        selecionar_funcionarios()
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    elif relatorio == 'Vale Transporte':
        # Aba Vale Transporte > Imprimir > Recibo de Vale Transporte Individual
        _click_img('ValeTransporte.png')
        _wait_img('AbaValeTransporte.png', conf=0.9, timeout=-1)
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('SelecionaValeTransporte.png', conf=0.9, timeout=-1)
        press(['r', 'enter'], interval=0.5)
        selecionar_funcionarios()
        _wait_img('DataDeEntrega.png', conf=0.9, timeout=-1)

        # Calcular o último dia util do mês para digitar na data de entrega do vale transporte
        date = datetime.now().replace(month=int(mes)) + relativedelta(day=31)
        dia = date.day
        if date.isoweekday() == 6:
            dia -= 1
        if date.isoweekday() == 7:
            dia -= 2
        write(str(dia))

        sleep(1)
        press(['tab', 'enter', 'enter'], interval=0.7)
        if not imprimir(relatorio, andamentos, empresa, texto, espera=30, diretorio=relatorio):
            return False

    else:
        # Imprimir > Impressão Multipla
        if not gerar(relatorio, andamentos, empresa):
            return False
        _wait_img('SelecionarImpressao.png', conf=0.9, timeout=-1)
        press(['i', 'i', 'enter'], interval=0.5)
        sleep(2)
        if relatorio == 'Relatório Geral - Impressão Multipla':
            hotkey('ctrl', 't')
        if relatorio == 'Resumo Geral Mensal - Impressão Multipla':
            _click_img('ResumoGeralMensal.png', conf=0.9, clicks=2)
        if relatorio == 'Resumo Mensal - Impressão Multipla':
            _click_img('ResumoMensal.png', conf=0.9, clicks=2)
        if relatorio == 'Folha de Pagamento - Impressão Multipla':
            _click_img('13Multipla.png', conf=0.9, clicks=2)
            _click_img('FolhaAnaliticaMultipla.png', conf=0.9, clicks=2)

        sleep(1)
        hotkey('ctrl', 'p')
        sleep(2)
        if _find_img('Continuar.png', conf=0.9):
            _click_img('Continuar.png', conf=0.9)
            press('s')
        sleep(4)
        imprimir(relatorio, andamentos, empresa, texto, espera=500, diretorio=relatorio)

    return True


def relatorio_dpcuca(index, empresas, relatorio, andamentos):
    if relatorio == 'Experiência' or relatorio == 'Empresas':
        # Entra no relatório informado pelo usuário
        if not relatoriozinhos(relatorio, andamentos, 'N'):
            return False
        _inicial('dpcuca')
        return 'OK'

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod, cnpj, nome, mes, ano = empresa

        # Verificar horário
        _hora_limite = datetime.now().replace(hour=18, minute=25, second=0, microsecond=0)
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('dpcuca')
            getWindowsWithTitle('DPCUCA')[0].maximize()

        _indice(count, total_empresas, empresa)

        _inicial('dpcuca')

        # Verificações de login
        if not _login(empresa, 'Codigo', 'dpcuca', relatorio, mes, ano):
            press('enter')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Empresa Inativa', nome=andamentos)
            print('❌ Empresa Inativa')
            continue

        # CNPJ com os separadores para poder verificar a empresa no cuca
        if not _verificar_empresa(cnpj, 'dpcuca'):
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Empresa não encontrada no DPCUCA', nome=andamentos)
            print('❌ Empresa não encontrada no DPCUCA')
            continue

        # Entra na opção de cálculos do DPCUCA
        evento, imagen, cont = calculos(relatorio)

        # verifica se tem funcionário cadastrado
        if _find_img('Sem{}.png'.format(imagen), conf=0.9):
            _click_img('Sair.png', conf=0.9)
            _inicial('DPCUCA')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Sem {evento} cadastrado', nome=andamentos)
            print('❌ Sem {} cadastrado'.format(evento))
            continue

        # Entra no relatório informado pelo usuário
        if not relatoriozinhos(relatorio, andamentos, empresa):
            continue

        _inicial('dpcuca')


@_time_execution
def run():
    relatorio = escolher_relatorio()
    empresas = _open_lista_dados()
    andamentos = relatorio
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    _iniciar('DPCUCA')
    if relatorio_dpcuca(index, empresas, relatorio, andamentos) == 'OK':
        print('Relatório de {} gerado.'.format(relatorio))
    _fechar()


if __name__ == '__main__':
    run()
