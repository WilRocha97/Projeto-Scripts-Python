# -*- coding: utf-8 -*-
import pyperclip, time, os, shutil, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start
from dominio_comum import _login, _salvar_pdf


def dirf(empresa, ano, andamento):
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('n')
    # federais
    time.sleep(0.5)
    p.press('f')
    # DIRF
    time.sleep(0.5)
    p.press('i')
    # apartir de 2010
    time.sleep(0.5)
    p.press('p')
    
    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)
    
    # digita o ano base
    p.write(ano)
    time.sleep(0.5)
    p.press('tab')
    
    # digita o código do responsável
    p.write('1')
    time.sleep(0.5)
    p.press('tab')
    
    erro = arquivos_dirf(empresa, ano, andamento)
    
    relatorio_dirf(empresa, ano, andamento, erro)
    
    p.press('esc', presses=4)
    time.sleep(2)

    
def relatorio_dirf(empresa, ano, andamento, erro):
    # seleciona para gerar o relatório
    if _find_img('formulario.png', conf=0.9):
        _click_img('formulario.png', conf=0.9)
    time.sleep(1)
    
    # seleciona para gerar com folha de pagamento
    if _find_img('folha_de_pagamento.png'):
        _click_img('folha_de_pagamento.png')
    time.sleep(1)
    
    # abre a janela de outros dados
    p.hotkey('alt', 'u')
    
    while not _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)

    if _find_img('folha_de_pagamento.png', conf=0.95):
        p.press('esc')
        time.sleep(1)
        p.hotkey('alt', 'n')
        time.sleep(1)
        p.press('esc')
        
        _escreve_relatorio_csv(f'Não é possível editar aba de folha de pagamento;{erro}', nome=andamento)
        print('❌ Não é possível editar aba de folha de pagamento')
        return False
    
    _click_img('tem_folha.png', conf=0.95)
    time.sleep(2)
    
    # clicar para gerar de todos os colaboradores
    if _find_img('todos.png', conf= 0.95):
        _click_img('todos.png', conf=0.95)
    time.sleep(1)
    
    if _find_img('gerar_info_complementar.png', conf= 0.99):
        _click_img('gerar_info_complementar.png', conf=0.99)
    time.sleep(1)
    
    if _find_img('limitar_60_caractere.png', conf=0.99):
        _click_img('limitar_60_caractere.png', conf=0.99)
    time.sleep(1)
    
    p.hotkey('alt', 'g')
    
    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'o')
    
    while not _find_img('relatorio_gerado.png', conf=0.9):
        time.sleep(1)

        if _find_img('relatorio_gerado_2.png', conf=0.9):
            _salvar_pdf()
            mover_relatorio_2(empresa)
            break

        if _find_img('sem_dados_arquivo.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(f'Sem dados para emitir;{erro}', nome=andamento)
            print('❗ Sem dados para emitir')
            return False

    if _find_img('relatorio_gerado.png', conf=0.9):
        salvar_pdf(empresa)
    
    _escreve_relatorio_csv(f'Relatório DIRF {ano} gerado;{erro}', nome=andamento)
    print('✔ Relatório gerado')

    return True


def arquivos_dirf(empresa, ano, andamento):
    cod, cnpj, nome = empresa
    erro = ''
    
    # seleciona para gerar o relatório
    if _find_img('arquivo.png', conf=0.9):
        _click_img('arquivo.png', conf=0.9)
    time.sleep(1)
    
    # desce para a linha onde sera digitado o caminho para salvar o arquivo
    p.press('tab')
    time.sleep(1)
    
    # apaga qualuer texto que esteja no campo
    p.press('del', presses=85)
    
    # digita o caminho para salvar o arquivo
    os.makedirs('execução/Arquivos', exist_ok=True)
    pyperclip.copy('')
    pyperclip.copy('V:\Setor Robô\Scripts Python\Domínio\Gera Arquivo e Relatório DIRF')
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    
    # seleciona para gerar com folha de pagamento
    if _find_img('folha_de_pagamento.png'):
        _click_img('folha_de_pagamento.png')
    time.sleep(1)
    
    # abre a janela de outros dados
    p.hotkey('alt', 'u')

    while not _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'g')

    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'o')
    
    while not _find_img('dirf_gerada.png', conf=0.9):
        time.sleep(1)
        if _find_img('outros_dados_nao_digitados.png', conf=0.9):
            erro = 'Outros dados não digitados'
            p.hotkey('enter')
        
        if _find_img('sem_dados_arquivo.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Sem dados para emitir', nome=andamento, end=';')
            print('❗ Sem dados para emitir')
            return erro

    _click_img('dirf_gerada.png', conf=0.9)
    p.press('enter')
    time.sleep(2)
    
    if _find_img('60_caracteres.png', conf=0.9):
        erro = 'Descrição de outros rendimentos isentos e não-tributáveis superior a 60 caracteres'
        p.hotkey('alt', 'f')
        
    _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Arquivo DIRF {ano} gerado', nome=andamento, end=';')
    print('✔ Arquivo gerado')

    mover_arquivo(empresa)
    return erro


def salvar_pdf(empresa):
    p.hotkey('ctrl', 'd')
    
    timer = 0
    while not _find_img('procurar_pasta.png', conf=0.9):
        time.sleep(1)
        if _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
            p.press('esc')
            _salvar_pdf()
            mover_relatorio_3(empresa)
            return True

        timer += 1
        if timer > 30:
            return False
    
    if not _find_img('cliente_c_selecionado.png', conf=0.9):
        while not _find_img('cliente_c.png', conf=0.9):
            _click_img('botao.png', conf=0.9)
            time.sleep(3)
        
        _click_img('cliente_c.png', conf=0.9, timeout=1)
        time.sleep(3)

    while not _find_img('arquivos.png', conf=0.9):
        time.sleep(1)

    _click_img('arquivos.png', conf=0.9, timeout=1)
    
    p.hotkey('alt', 'o')
    
    while not _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('adobe.png', pasta='imgs_c', conf=0.9):
            _click_img('adobe.png', pasta='imgs_c', conf=0.9)
    
    while _find_img('pdf_aberto.png', pasta='imgs_c', conf=0.9):
        p.hotkey('alt', 'f4')
        time.sleep(3)
    
    while _find_img('sera_finalizada.png', pasta='imgs_c', conf=0.9):
        p.press('esc')

    mover_relatorio(empresa)
    return True


def mover_arquivo(empresa):
    cod, cnpj, nome = empresa

    download_folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF"
    folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF\\execução\\Arquivos"
    guia = os.path.join(download_folder, f'DIRF{cod}.txt')
    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(folder, f'DIRF{cod}.txt'))
            time.sleep(2)
        except:
            pass


def mover_relatorio(empresa):
    cod, cnpj, nome = empresa
    cnpj = cnpj.replace('/', '').replace('.', '').replace('-', '')
    nome = nome.replace('/', ' ').replace('?', ' ').replace(':', ' ').replace('"', ' ').replace('*', ' ')
    os.makedirs('execução/Relatórios', exist_ok=True)

    download_folder = "C:\\Arquivos"
    folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF\\execução\\Relatórios"
    guia = os.path.join(download_folder, 'DIRF.pdf')
    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(folder, f'DIRF{cod} - {cnpj} - {nome}.pdf'))
            time.sleep(2)
        except:
            pass


def mover_relatorio_2(empresa):
    cod, cnpj, nome = empresa
    cnpj = cnpj.replace('/', '').replace('.', '').replace('-', '')
    nome = nome.replace('/', ' ')
    os.makedirs('execução/Relatórios', exist_ok=True)

    folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF\\execução\\Relatórios"
    guia = os.path.join('C:\\DIRF Escrita.pdf')
    while not os.path.exists(guia):
        time.sleep(1)
        
    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(folder, f'DIRF{cod} - {cnpj} - {nome}.pdf'))
            time.sleep(2)
        except:
            pass



def mover_relatorio_3(empresa):
    cod, cnpj, nome = empresa
    cnpj = cnpj.replace('/', '').replace('.', '').replace('-', '')
    nome = nome.replace('/', ' ')
    os.makedirs('execução/Relatórios', exist_ok=True)

    folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF\\execução\\Relatórios"
    guia = os.path.join('C:\\DIRF Folha.pdf')
    while not os.path.exists(guia):
        time.sleep(1)

    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(folder, f'DIRF{cod} - {cnpj} - {nome}.pdf'))
            time.sleep(2)
        except:
            pass


@_time_execution
def run():
    ano = p.prompt(text='Qual ano base?', title='Script incrível', default='0000')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos DIRF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
    
        if not _login(empresa, andamentos):
            continue
        dirf(empresa, ano, andamentos)


if __name__ == '__main__':
    run()
