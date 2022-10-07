# -*- coding: utf-8 -*-
import sys
import os
from time import sleep
from pyperclip import copy
from pyautogui import press, write, hotkey

sys.path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _download_file, _open_lista_dados, _where_to_start, _indice
from pyautogui_comum import _find_img, _click_img, _wait_img


def salvar_pdf(cnpj, nome, debito=''):
    # navega na tela até aparecer o botão de emitir relatório
    while not _find_img('emitir_relatorios.png', conf=0.9):
        press('pgDn')
    _click_img('emitir_relatorios.png', conf=0.9)
    _wait_img('salvar.png', conf=0.9, timeout=-1)
    sleep(1)
    # escreve o nome do PDF
    copy(f'{nome} - {cnpj} - Certidão Negativa de Débitos Não Inscritos{debito}.pdf')
    hotkey('ctrl', 'v')
    sleep(1)
    # vai até o campo para inserir o caminho que irá salvar o PDF
    press('tab', presses=6)
    sleep(0.5)
    press('enter')
    sleep(0.5)

    # cola o caminho que vai salvar o PDF com o pyperclip porque o pyautogui não consegue copiar e colar texto com acentuação
    docs = 'V:\Setor Robô\Scripts Python\Fazenda\Consulta de Certidão Negativa de Débitos Tributários Não Inscritos\execucao\docs'
    os.makedirs(r'{}'.format(docs), exist_ok=True)
    copy(docs)
    hotkey('ctrl', 'v')

    # salva o PDF
    sleep(0.5)
    press('enter')
    sleep(0.5)
    hotkey('alt', 'l')
    sleep(1)

    # caso já exista um PDF com o mesmo nome ele substitui
    if _find_img('salvar_como.png', conf=0.9):
        press('s')
        sleep(1)

    texto = f'{cnpj};Com pendências {debito}'
    print(f'❗ Com pendências {debito}')
    _escreve_relatorio_csv(texto)

    # esperar aparecer o botão de voltar e clica nele
    _wait_img('voltar.png', conf=0.9)
    _click_img('voltar.png', conf=0.9)


def consulta_ipva(cnpj, nome):
    # url para entrar no site
    # url = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages/Restrita/PesquisarContribuinte.aspx'

    # espera a pagina inicial para inserir o cnpj
    while not _find_img('cnpj.png', conf=0.9):
        if _find_img('certificado.png', conf=0.9):
            _click_img('certificado.png', conf=0.9)
        if _find_img('verificar_impedimentos.png', conf=0.9):
            _click_img('verificar_impedimentos.png', conf=0.9)
        sleep(1)

    _click_img('campo.png', conf=0.9)
    sleep(1)
    # limpa o campo do cnpj
    press('delete', presses=15)
    write(cnpj)
    sleep(1)
    _click_img('consultar.png', conf=0.9)
    sleep(2)

    # aguarda a tela de carregamento
    while _find_img('aguarde.png', conf=0.9):
        sleep(1)

    # espera a tela da empresa abrir e caso apareça a tela de erro da F5 na página
    while not _find_img('dados.png', conf=0.9):
        if _find_img('atencao.png', conf=0.9):
            press('enter')
            press('f5')

    debitos = ('ha_debitos.png', 'ha_pendencias.png', 'ha_pendencias_2.png', 'icms_declarado.png', 'icms_parcelado.png', 'aiim.png', 'ipva.png')
    # navega na tela até aparecer o botão de emitir relatório e caso tenha algum débito, salva o relatório
    while not _find_img('emitir_relatorios.png', conf=0.9):
        press('pgDn')
        
        for debito in debitos:
            if _find_img(debito, conf=0.9):
                if _find_img('gia.png', conf=0.9):
                    salvar_pdf(cnpj, nome, debito=' - GIA')
                else:
                    salvar_pdf(cnpj, nome)
                return True

    # define o texto que ira escrever na planilha
    texto = f'{cnpj};Empresa sem pendências'
    print('✔ Empresa sem pendências')
    _escreve_relatorio_csv(texto)

    # voltar pra tela de login da empresa
    _wait_img('voltar.png', conf=0.9)
    _click_img('voltar.png', conf=0.9)
    return False


@_time_execution
def run():
    # abre arquivo de dados
    empresas = _open_lista_dados()

    # define a primeira empresa que vai executar o script
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # define o número total de empresas
    total_empresas = empresas[index:]

    # começa a repetição
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        nome = nome.replace('/', ' ')
        # printa o indice dos dados
        _indice(count, total_empresas, empresa)

        # executa a consulta
        consulta_ipva(cnpj, nome)


if __name__ == '__main__':
    run()
