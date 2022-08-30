# -*- coding: utf-8 -*-
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tkinter import *
from ttk import Combobox
import time, pyperclip, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def calcula_dia(mes_final):
    date = datetime.now().replace(month=int(mes_final)) + relativedelta(day=31)
    dia = date.day
    return dia


def configura_automatico(sped):
    consulta, ano_consulta, mes, mes_final, dia = '', '', '', '', ''
    
    if sped == 'SPED Contribuições':
        consulta = 'Consulta SPED Contribuições'

        mes = datetime.now().month
        if mes == 1:
            mes = 11
        elif mes == 2:
            mes = 12
        else:
            mes -= 2

        ano_consulta = datetime.now().strftime('%Y')

        if mes >= 11:
            ano_consulta = int(ano_consulta) - 1
        else:
            if mes <= 9:
                mes = str(mes)
                mes = '0' + mes

        mes_final = mes
        dia = calcula_dia(mes_final)

    elif sped == 'SPED Fiscal':
        consulta = 'Consulta SPED fiscal'
        mes = datetime.now().month

        if mes == 1:
            mes = 12
        else:
            mes -= 1

        ano_consulta = datetime.now().strftime('%Y')

        if mes == 12:
            ano_consulta = int(ano_consulta) - 1
        else:
            mes = int(mes)
            if mes <= 9:
                mes = str(mes)
                mes = '0' + mes

        mes_final = mes
        dia = calcula_dia(mes_final)

    elif sped == 'SPED ECF':
        consulta = 'Consulta SPED ECF'
        ano_consulta = datetime.now().strftime('%Y')
        mes = '01'
        mes_final = datetime.now().strftime('%m')
        dia = datetime.now().strftime('%d')
    
    elif sped == 'SPED ECD':
        consulta = 'Consulta SPED ECD'
        ano_consulta = int(datetime.now().strftime('%Y')) - 1
        mes = '01'
        mes_final = '12'
        dia = '31'

    return consulta, sped, str(ano_consulta), str(mes), str(mes_final), str(dia)


def configura_manual(sped):
    consulta, ano_consulta, mes, mes_final, dia = '', '', '', '', ''
    
    if sped == 'SPED Contribuições':
        consulta = 'Consulta SPED Contribuições'
        mes, mes_final, ano_consulta = pergunta_competencia()
        dia = calcula_dia(mes_final)

    elif sped == 'SPED Fiscal':
        consulta = 'Consulta SPED Fiscal'
        mes, mes_final, ano_consulta = pergunta_competencia()
        dia = calcula_dia(mes_final)

    elif sped == 'SPED ECF':
        consulta = 'Consulta SPED ECF'
        ano_consulta = p.prompt(title='Script incrível', text='Qual o período de entrega?', default='0000')
        mes = '01'
        if str(datetime.now().year) == ano_consulta:
            mes_final = datetime.now().strftime('%m')
            dia = datetime.now().strftime('%d')
        else:
            mes_final = '12'
            dia = calcula_dia(mes_final)
    
    elif sped == 'SPED ECD':
        consulta = 'Consulta SPED ECD'
        ano_consulta = p.prompt(title='Script incrível', text='Qual o período da escrituração?', default='0000')
        mes = '01'
        mes_final = '12'
        dia = '31'

    return consulta, sped, ano_consulta, mes, mes_final, dia


def pergunta_competencia():
    competencia = p.prompt(title='Script incrível', text='Qual o período da escrituração?', default='00/0000')
    competencia = competencia.split('/')
    mes = competencia[0]
    mes_final = competencia[0]
    ano_consulta = competencia[1]

    return mes, mes_final, ano_consulta


def abrir_receita(empresa):
    cnpj, nome, cert = empresa
    time.sleep(0.5)
    if _find_img('Atalho.png', conf=0.9):
        _click_img('Atalho.png', conf=0.9, clicks=2)
    else:
        _click_img('Atalho2.png', conf=0.9, clicks=2)
    _wait_img('Receitanet.png', conf=0.9, timeout=-1)
    while not _find_img('Cert' + cert + '.png', conf=0.9):
        _click_img('down.png', conf=0.9)
    _click_img('Cert' + cert + '.png', conf=0.9)

    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.write('19937946000107')
    time.sleep(1)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    _click_img('Maximiza.png', conf=0.9)
    time.sleep(1)


def fechar_receita():
    if _find_img('Icone.png', conf=0.9):
        p.getWindowsWithTitle("Receitanet BX")[0].maximize()
        time.sleep(0.5)
    p.getWindowsWithTitle('Receitanet BX')[0].close()
    time.sleep(0.5)


def trocar_empresa(cnpj):
    # Trocar a empresa que será pesquisada
    while not _find_img('Receitanet.png', conf=0.9):
        p.hotkey('ctrl', 't')
        time.sleep(0.5)
        
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('backspace', presses=14)
    time.sleep(0.5)
    p.write(cnpj)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)


def verificacoes(empresa, consulta):
    cnpj_form, nome, cert = empresa
    cnpj = cnpj_form.replace('/', '').replace('-', '').replace('.', '')

    while _find_img('NaoConcluida.png', conf=0.9):
        if _find_img('Revogado.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'Certificado revogado']), nome=consulta)
            response = '❌ Certificado revogado'
            return response
        p.press('enter')
        time.sleep(3)
        p.hotkey('ctrl', 'p')
        time.sleep(3)

    if _find_img('ProcRejeitada.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'A procuração eletrônica está pendente de aprovação']), nome=consulta)
        response = '❌ A procuração eletrônica está pendente de aprovação'

    elif _find_img('Pendente.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'A procuração eletrônica foi cancelada']), nome=consulta)
        response = '❌ A procuração eletrônica foi cancelada'

    elif _find_img('ProcCancelada.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'A procuração eletrônica foi cancelada']), nome=consulta)
        response = '❌ A procuração eletrônica foi cancelada'

    elif _find_img('ProcExpirou.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'A procuração eletrônica expirou']), nome=consulta)
        response = '❌ A procuração eletrônica expirou'

    elif _find_img('NaoExisteProc.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'Não existe procuração eletrônica para o detentor do certificado digital apresentado']), nome=consulta)
        response = '❌ Não existe procuração eletrônica para o detentor do certificado digital apresentado'

    elif _find_img('NaoArquivo.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'Nenhum arquivo encontrado']), nome=consulta)
        response = '❌ Nenhum arquivo encontrado'

    elif _find_img('NaoArquivo2.png', conf=0.9):
        p.press('enter')
        _escreve_relatorio_csv(';'.join([cnpj_form, cert, nome, cnpj, 'Nenhum arquivo encontrado']), nome=consulta)
        response = '❌ Nenhum arquivo encontrado'

    else:
        response = 'ok'

    return response


def configurar_contribuicoes():
    # Configurar pesquisa contribuições
    _wait_img('SelecionarSistema.png', conf=0.9, timeout=-1)
    _click_img('SelecionarSistema.png', conf=0.9)
    p.press('esc')
    time.sleep(0.5)
    p.moveTo(1150, 200)
    p.press('down')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    _click_img('TipoDePesquisa.png', conf=0.9)
    _wait_img('PeriodoDeEscrituracao.png', conf=0.9, timeout=-1)
    _click_img('PeriodoDeEscrituracao.png', conf=0.9)


def configurar_fiscal():
    # Configurar pesquisa fiscal
    _wait_img('SelecionarSistema.png', conf=0.9, timeout=-1)
    _click_img('SelecionarSistema.png', conf=0.9)
    p.press('esc')
    time.sleep(0.5)
    p.moveTo(1150, 200)
    p.press('down', presses=5)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press(['tab', 'tab', 'tab'], interval=0.5)


def configurar_ecf():
    # Configurar pesquisa ecf
    _wait_img('SelecionarSistema.png', conf=0.9, timeout=-1)
    _click_img('SelecionarSistema.png', conf=0.9)
    p.press('esc')
    time.sleep(0.5)
    p.moveTo(1150, 200)
    p.press('down', presses=3)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')


def configurar_ecd():
    # Configurar pesquisa ecd
    _wait_img('SelecionarSistema.png', conf=0.9, timeout=-1)
    _click_img('SelecionarSistema.png', conf=0.9)
    p.press('esc')
    time.sleep(0.5)
    p.moveTo(1150, 200)
    p.press('down', presses=2)
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('down')


def consulta_sped(empresas, index, consulta, sped, ano_consulta, mes, mes_final, dia):
    p.hotkey('win', 'm')

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):

        _indice(count, total_empresas, empresa)

        if not _find_img('Pesquisa.png', conf=0.9):
            if not _find_img('Icone.png', conf=0.9):
                abrir_receita(empresa)
            else:
                p.getWindowsWithTitle("Receitanet BX")[0].maximize()

        _click_img('Pesquisa.png', conf=0.9)
        if count == 1:
            if not _find_img('SelecionarSistema2.png', conf=0.9):
                trocar_empresa('19937946000107')

        cnpj_form, nome, cert = empresa
        cnpj = cnpj_form.replace('/', '').replace('-', '').replace('.', '')

        # Foca no programa da Receitanet BX e reinicia o programa caso o procurador estiver errado
        _click_img('Pesquisa.png', conf=0.9, delay=0)
        if not _find_img('Usuario' + cert + '.png', conf=0.9):
            fechar_receita()
            abrir_receita(empresa)

        time.sleep(1)

        # Trocar a empresa que será pesquisada
        trocar_empresa(cnpj)

        if sped == 'SPED Contribuições':
            configurar_contribuicoes()
        if sped == 'SPED Fiscal':
            configurar_fiscal()
        if sped == 'SPED ECF':
            configurar_ecf()
        if sped == 'SPED ECD':
            configurar_ecd()

        # Digitar data inicial (Os dias devem ter um delay entre um dígito e outro para não dar problema no programa da Receitanet)
        time.sleep(0.5)
        p.write('01', interval=0.5)
        time.sleep(0.5)
        p.write(mes)
        p.write(ano_consulta)
        time.sleep(0.5)
        p.press('tab')

        # Digitar data final (Os dias devem ter um delay entre um dígito e outro para não dar problema no programa da Receitanet)
        p.write(str(dia), interval=0.5)
        time.sleep(0.5)
        p.write(mes_final)
        p.write(ano_consulta)
        time.sleep(1)
        p.press('tab')

        p.hotkey('ctrl', 'p')
        time.sleep(1)

        while _find_img('Pesquisando.png', conf=0.9):
            time.sleep(2)

        while not _find_img('ResultadosPesquisa.png', conf=0.9):
            time.sleep(1)
            response = verificacoes(empresa, consulta)
            if response != 'ok':
                print(response)
                break

        while _find_img('Resultado.png', conf=0.99):
            _click_img('Resultado.png', conf=0.99)
            p.hotkey('ctrl', 'c')
            resultado = pyperclip.paste()
            resultado = resultado.split()
            resultado = str(resultado)
            resultado = resultado.replace('true', '{};{};{}'.format(cnpj_form, cert, nome)).replace("', '", ';').replace("['", '').replace("']", '')
            _escreve_relatorio_csv(texto=resultado, nome=consulta)
            print('✔ Arquivo encontrado')

    p.getWindowsWithTitle('Receitanet BX')[0].close()


def qual_consulta():
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
    
    relatorios = ["SPED Contribuições", "SPED Fiscal", "SPED ECF", "SPED ECD"]
    
    # drop down menu
    Combobox(root, textvariable=relatorio, values=relatorios, width=55).pack()
    
    # botão com comando para fechar o tkinter
    Button(root, text="Ok", command=root.destroy).pack()
    
    root.mainloop()
    
    return str(relatorio.get())


def consultar(sped, empresas, index):
    manual = p.confirm(title='Script incrível', text='informar competência específica?', buttons=['Sim', 'Não'])

    if manual == 'Sim':
        consulta, sped, ano_consulta, mes, mes_final, dia = configura_manual(sped)

    else:
        consulta, sped, ano_consulta, mes, mes_final, dia = configura_automatico(sped)

    consulta_sped(empresas, index, consulta, sped, ano_consulta, mes, mes_final, dia)


@_time_execution
def run():
    sped = qual_consulta()
    
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    consultar(sped, empresas, index)


if __name__ == '__main__':
    run()
