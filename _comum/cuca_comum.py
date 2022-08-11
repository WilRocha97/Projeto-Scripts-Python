# -*- coding: utf-8 -*-
from datetime import datetime
from comum_comum import _escreve_relatorio_csv
import pyautogui as p
import time, sys


def _loc_img(imagem, conf=0.9):
    if p.locateOnScreen(r'imgs_c\{}'.format(imagem), confidence=conf):
        return True
    else:
        return False


def _horario(_hora_limite, qual_cuca):
    # Verifica o horário
    if datetime.now() >= _hora_limite:
        _inicial(qual_cuca)
        _fechar()
        voltar = p.confirm('Posso continuar?', buttons=['Continuar', 'Parar robô'])
        if voltar == 'Continuar':
            return True
        else:
            sys.exit()


def esperar(imagem):
    while not p.locateOnScreen(r'imgs_c\{}'.format(imagem), confidence=0.9):
        time.sleep(0.5)


def clicar(imagem, cliques=1, conf=0.9):
    try:
        if _loc_img(imagem):
            if cliques == 2:
                p.doubleClick(p.locateCenterOnScreen(r'imgs_c\{}'.format(imagem), confidence=conf))
            else:
                p.click(p.locateCenterOnScreen(r'imgs_c\{}'.format(imagem), confidence=conf))
        return True
    except:
        return False


def _iniciar(qual_cuca):
    cuca = {'CUCA': '(SP)', 'DPCUCA': 'DPCUCA'}
    if not _loc_img(qual_cuca + '.png'):
        if not _loc_img(qual_cuca + '2.png'):
            iniciar_atalho()
            escolher(qual_cuca)
    else:
        p.hotkey('win', 'm')
        time.sleep(0.5)
        p.getWindowsWithTitle(cuca[qual_cuca])[0].maximize()
        time.sleep(1)
        clicar('T' + qual_cuca + 'Inicial.png')
        time.sleep(5)


def _inicial(qual_cuca):
    p.getWindowsWithTitle(qual_cuca)[0].maximize()
    time.sleep(0.5)
    while _loc_img('JanelasDo' + qual_cuca + '.png'):
        clicar('JanelasDo' + qual_cuca + '.png')
        p.hotkey('alt', 'f4')
        time.sleep(0.5)
        p.press(['esc'], presses=2, interval=0.5)
        time.sleep(0.5)


def iniciar_atalho():
    # Verifica se o atalho ja está aberto, se estiver só maximiza
    if _loc_img('AtalhoCUCA.png'):
        clicar('AtalhoCUCA.png')
    elif _loc_img('AtalhoCUCA2.png'):
        clicar('AtalhoCUCA2.png')
    else:
        usuario = p.prompt(text='Qual o usuario do robô?', title='Script incrível')
        p.hotkey('win', 'm')
        time.sleep(0.5)
        clicar('Atalho.png' or
               'Atalho2.png' or
               'Atalho3.png', cliques=2)

        senhas = {'robo': 'RbCuca#25', 'robo2': 'robo#2', 'robo3': 'robo#3', 'robo4': 'robo#4', 'robo5': 'robo#5', 'robo6': 'robo#6', }
        acesso = 'negado'
        while acesso == 'negado':

            senha = (senhas[usuario])
            esperar('LoginCUCA.png')
            # Login no atalho do CUCA
            clicar('LoginCUCA.png')
            p.write(senha)
            p.press('tab')
            p.write(usuario)
            p.press(['tab', 'enter'], interval=0.5)
            time.sleep(5)
            if _loc_img('AcessoNegado.png'):
                p.press('enter')
                usuario = p.prompt(text='Qual o usuario do robô?', title='Script incrível')
                acesso = 'negado'
            else:
                acesso = 'validado'


def escolher(qual_cuca):
    # Seleciona qual cuca deve ser aberto
    while not _loc_img('BotaoIniciaAtalho.png'):
        clicar('AvisoAtalho.png')
    clicar('JanelaDeInicio.png')
    clicar('Botao' + qual_cuca + '.png')
    clicar('BotaoIniciaAtalho.png')
    time.sleep(2)
    if _loc_img('AlertaAtalho.png'):
        p.hotkey('alt', 'o')
    esperar('Escolher' + qual_cuca + '.png')
    p.press('enter')
    time.sleep(2)
    _verificacoes(qual_cuca, '1')
    time.sleep(2)


def _login(empresa, log, qual_cuca, execucao='cuca', competencia=str(datetime.now().month), ano=str(datetime.now().year)):
    try:
        cod, nome, cnpj = empresa
    except:
        cod, nome, cnpj = '0000', 'EMPRESA', '00.000.000/0000-00'
    # Loga em uma empresa
    cuca = {'cuca': ['o', 'j', '(SP)'], 'dpcuca': ['g', 'p', 'DPCUCA']}

    # Maximizar o cuca e verifica se está na página inicial para abrir a janela de seleção de empresa
    while not _loc_img('referencia' + qual_cuca + '.png'):
        p.getWindowsWithTitle(cuca[qual_cuca][2])[0].maximize()

    time.sleep(1)
    p.press('esc')

    while not _loc_img('Escolher' + qual_cuca + '.png'):
        clicar('T' + qual_cuca + 'Inicial.png')
        clicar('referencia' + qual_cuca + '.png')
        time.sleep(2)

    # Escolher a empresa e a competência

    clicar('Escolher' + qual_cuca + '.png')
    competencia = competencia.replace('12', 'c').replace('11', 'b').replace('10', 'a')
    p.hotkey('alt', competencia)
    time.sleep(1)
    # Se for o DPCUCA seleciona o ano
    if qual_cuca == 'dpcuca':
        p.hotkey('shift', 'tab')
        time.sleep(1)
        p.write(ano)

    if not empresa == 'N':
        # Seleciona a aba para pesquisar a empresa por código ou CNPJ
        dicionario = {'Codigo': [empresa[0], cuca[qual_cuca][0]], 'CNPJ': [empresa[1], cuca[qual_cuca][1]]}
        p.hotkey('alt', dicionario[log][1])
        time.sleep(1)

        clicar('EscolherEmpresa' + qual_cuca + '.png')
        p.write(dicionario[log][0])
        time.sleep(1)
        if _loc_img('EmpresaRecibo.png'):
            p.press('esc')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Empresa Recibo']), nome=execucao)
            print('❌ Empresa Recibo')
            return False

    p.press('enter')

    while _loc_img('Escolher' + qual_cuca + '.png'):
        if _loc_img('SenhaDoMes.png'):
            clicar('SenhaDoMes.png')
            p.press('enter')
            time.sleep(2)
        time.sleep(0.2)

    time.sleep(10)
    if not _verificacoes(qual_cuca, competencia, execucao, empresa):
        return False

    clicar('T' + qual_cuca + 'Inicial.png')
    return True


def _verificacoes(qual_cuca, competencia, execucao='cuca', empresa='N'):
    try:
        cod, nome, cnpj = empresa
    except:
        cod, nome, cnpj = '0000', 'EMPRESA', '00.000.000/0000-00'
    # Tratar as mensagens aleatórias que aparecer
    if _loc_img('CodContabeis.png'):
        clicar('AplicarCodContabeis.png')
    if _loc_img('EmpresaNaoAtivada.png'):
        clicar('sair.png')
        return False

    if qual_cuca == 'cuca':
        if competencia == '01' or competencia == '1':
            time.sleep(12)
            while _loc_img('ImportarCadastro.png'):
                clicar('ImportarCadastro.png')
                p.hotkey('alt', 's')
                time.sleep(1)
                while _loc_img('AtualizandoSistema.png'):
                    time.sleep(1)
                time.sleep(15)
                while _loc_img('NaoResponde.png'):
                    time.sleep(1)
                time.sleep(15)
                if _loc_img('NaoImportouCadastro.png'):
                    p.press('enter')
                    clicar('EscolherEmpresa' + qual_cuca + '.png')
                    p.press('enter')
                p.press('enter')
        if _loc_img('NaoSocio.png'):
            clicar('NaoSocio.png')
            p.hotkey('alt', 'o')
            time.sleep(2)
            p.press('esc')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Não é possível entrar na empresa não existe sócio cadastrado']), nome=execucao)
            print('❌ Não é possível entrar na empresa não existe sócio cadastrado')
            return False
        if _loc_img('OpcaoInvalida.png'):
            clicar('OK4.png')
            while not _loc_img('CucaWeb.png'):
                try:
                    p.getWindowsWithTitle('Chrome')[0].maximize()
                except:
                    pass
                if _loc_img('CucaWeb2.png'):
                    break
                time.sleep(2)
            p.hotkey('ctrl', 'w')
            time.sleep(5)
            clicar('EscolherEmpresaCUCA.png')
            time.sleep(0.5)
            p.press('esc')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Acesso bloqueado']), nome=execucao)
            print('❌ Acesso bloqueado')
            return False
        if _loc_img('EmpresaNaoLiberada.png'):
            clicar('OK4.png')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Não é possível entrar na empresa não liberada no período selecionado!']), nome=execucao)
            print('❌ Não é possível entrar na empresa não liberada no período selecionado!')
            p.press('esc')
            return False
        if _loc_img('EmpresaImpossivel.png'):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Não é possível entrar na empresa']), nome=execucao)
            print('❌ Não é possível entrar na empresa')
            time.sleep(2)
            return False
        if _loc_img('Reprocessar.png'):
            clicar('Reprocessar.png')
            p.hotkey('alt', 'n')
        if _loc_img('BalancoFechado.png'):
            p.press('enter')
        if _loc_img('Classificar.png'):
            p.press('esc')
        if _loc_img('AceitarAlteracao.png'):
            clicar('AceitarAlteracao.png')
        if _loc_img('0Socios.png'):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, nome, cnpj, 'Não é possível entrar na empresa, não existe sócio cadastrado']), nome=execucao)
            print('❌ Não é possível entrar na empresa, não existe sócio cadastrado')
            time.sleep(2)
            return False
        if _loc_img('Atencao.png'):
            p.press('enter')
            time.sleep(2)
        if _loc_img('Criticas.png'):
            p.press('enter')
            time.sleep(3)
            clicar('T' + qual_cuca + 'Inicial.png')

    elif qual_cuca == 'dpcuca':
        if _loc_img('TabelasDPCUCA.png'):
            clicar('Confirma.png', 2)
            time.sleep(2)
        if _loc_img('AtualizacaoESOCIAL.png'):
            clicar('AtualizacaoESOCIAL.png', 2)
            time.sleep(2)
            p.hotkey('alt', 'c')
            time.sleep(10)
        if _loc_img('Alerta.png'):
            p.press(['esc', 'enter'], interval=0.5)
            time.sleep(2)
        if _loc_img('FuncionarioSimulado.png'):
            clicar('FuncionarioSimulado.png')
            p.hotkey('alt', 'n')
        if _loc_img('OK2.png'):
            clicar('OK2.png')
            clicar('OK3.png')
            time.sleep(2)
        p.hotkey('alt', 'n')
        p.press('esc')
        if _loc_img('EmpresaInativa.png'):
            clicar('Sair.png')
            return False

    return True


def _verificar_empresa(indice, andamentos, texto='falso', qual_cuca='falso'):
    # Verifica se está logado na empresa certa
    try:
        if qual_cuca == 'cuca':
            while not _loc_img('Destravado.png'):
                time.sleep(1)
        empresa = p.getWindowsWithTitle(indice)[0].isMaximized
        if empresa:
            return True
    except:
        if texto != 'falso':
            _escreve_relatorio_csv(texto, nome=andamentos)
            print('❌ Empresa não encontrada no {}'.format(qual_cuca.upper()))
        return False


def _fechar():
    # Fecha os dois cucas se estiverem abertos
    tipos = [("(SP)", "CUCA"), ("DPCUCA PLUS", "DPCUCA"), ]
    p.hotkey('win', 'm')
    for tipo in tipos:
        time.sleep(1)
        if _loc_img(tipo[1] + '.png', conf=0.9):
            time.sleep(0.5)
            p.getWindowsWithTitle(tipo[0])[0].maximize()
            time.sleep(0.5)
            while not _loc_img('T' + tipo[1] + 'Inicial.png'):
                time.sleep(0.5)
                if not clicar('JanelasDoDPCUCA.png'):
                    clicar('JanelasDoCUCA.png')
                p.hotkey('alt', 'f4')
                time.sleep(0.5)
                p.press(['esc'], presses=2, interval=0.5)
                time.sleep(0.5)
            p.hotkey('alt', 'f4')
            esperar('Desligar' + tipo[1] + '.png')
            p.hotkey('alt', 'n')
            time.sleep(1)
    fechar_atalho()


def fechar_atalho():
    # Fecha o atalho do CUCA
    if _loc_img('AtalhoCUCA.png'):
        clicar('AtalhoCUCA.png')
    elif _loc_img('AtalhoCUCA2.png'):
        clicar('AtalhoCUCA2.png')

    clicar('AvisoAtalho.png')
    clicar('FecharAtalho.png')
    clicar('FecharAtalho.png', 1)
    if _loc_img('RecordErroAtalho.png'):
        p.hotkey('alt', 'o')
        time.sleep(0.5)
    clicar('OKAtalho.png')
    esperar('DesligarAtalho.png')
    clicar('DesligarAtalho.png')
