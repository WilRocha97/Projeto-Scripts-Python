# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p
from pyperclip import copy
import os

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def reinstala_sedif():
    time.sleep(2)
    p.hotkey('win', 'm')
    os.startfile(r'C:\Users\robo\Downloads\sedif_Setup')
    _wait_img('InstalaSedif.png', conf=0.9, timeout=-1)
    p.press('enter', presses=5, interval=1)
    _wait_img('FinalInstalaSedif.png', conf=0.9, timeout=-1)
    p.press('enter')
    _wait_img('AtualizaSedif.png', conf=0.9, timeout=-1)
    _click_img('AtualizaSedif.png', conf=0.9)
    time.sleep(1)
    _click_img('Nao.png', conf=0.9)
    time.sleep(5)
    if _find_img('AtualizaSedif2.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        p.press('esc')
    time.sleep(1)
    

def configura(empresa, comp):
    cnpj, nome = empresa
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]
    p.moveTo(100, 100)
    
    _wait_img('novo_documento.png', conf=0.9, timeout=-1)
    _click_img('novo_documento.png', conf=0.9)
    
    while not _find_img('NomeEmpresa.png', conf=0.9):
        time.sleep(10)
        p.click(900, 900)
    _click_img('NomeEmpresa.png', conf=0.9)
    
    p.write(nome)
    time.sleep(5)
    p.hotkey('ctrl', 'a')
    p.hotkey('ctrl', 'c')
    text = pyperclip.paste()
    print(text)
    
    if text != nome:
        _escreve_relatorio_csv(f'{cnpj};{nome};{text};Empresa não encontrada no SEDIF', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
        print('❗ Empresa não encontrada')
        p.press('esc')
        p.hotkey('alt', 'l')
        _wait_img('Novo.png', conf=0.9, timeout=-1)
        p.hotkey('alt', 'f')
        time.sleep(1)
        p.hotkey('alt', 'f4')
        time.sleep(1)
        p.hotkey('alt', 's')
        return False
    
    p.press('enter')
    p.write(comp)
    p.hotkey('alt', 'p')
    
    while not p.locateOnScreen(os.path.join('imgs', 'Escrituracao.png')):
        time.sleep(1)
    p.write('DeSTDA')
    p.press('tab', presses=2)
    p.write('Original')
    p.press('tab')
    p.write('Sem dados informados')
    time.sleep(0.5)
    _click_img('Confirmar.png', conf=0.9)
    time.sleep(20)
    
    if _find_img('jaExisteDeclaracao.png', conf=0.9):
        p.press('esc')
    
    if _find_img('PreenchimentoErrado.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};{nome};Preenchimento incorreto dos campos {comp}', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
        print('❗ Preenchimento incorreto dos campos')
        p.press('enter')
        p.hotkey('alt', 'l')
        time.sleep(2)
        p.hotkey('alt', 'f')
        return False
    
    p.hotkey('alt', 'f')
    return True


def excluir_documento():
    while not _find_img('abrir_arquivo.png', conf=0.9):
        time.sleep(1)
        p.press('up')
        _click_img('minimizar.png', conf=0.9)
    
    time.sleep(1)
    p.hotkey('alt', 'e')
    time.sleep(1)
    p.hotkey('alt', 's')
    
    # espera exlcuir o arquivo
    _wait_img('arquivo_excluido.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'f')
    

def transmitir(empresa, comp):
    # p.alert(text='alerta')
    cnpj, nome = empresa
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]

    p.hotkey('alt', 'f')
    
    # abrir arquivo da empresa
    while not _find_img('abrir_arquivo.png', conf=0.9):
        _wait_img('abrir_documento.png', conf=0.9, timeout=-1)
        if _find_img('abrir_documento.png', conf=0.9):
            _click_img('abrir_documento.png', conf=0.9)
        time.sleep(3)
        p.press('up')
        _click_img('minimizar.png', conf=0.9)
    p.press('enter')
    time.sleep(2)
    p.hotkey('alt', 'b')
    
    # espera tela inicial do sedif
    _wait_img('tela_inicial.png', conf=0.9, timeout=-1)
    
    # clicar em na aba encerrar
    _click_img('encerrar.png', conf=0.9)
    time.sleep(2)
    
    # gerar documento
    p.hotkey('alt', 'g')
    
    # esperar o botão para gerar assinatura do documento
   
    _wait_img('assinatura.png', conf=0.9)
    time.sleep(1)
    
    p.hotkey('alt', 'i')

    # espera o botão de transmitir
    while not _find_img('transmitir.png', conf=0.9):
        if _find_img('Criticas.png', conf=0.9):
            p.press('enter', presses=4, interval=2)
            time.sleep(1)
            p.hotkey('alt', 'f')
            time.sleep(1)
            _click_img('iniciar.png', conf=0.9)
            p.hotkey('alt', 'f')
            time.sleep(1)
            p.press('f')
            _escreve_relatorio_csv(f'{cnpj};{nome};Existe criticas nos dados que impedem a assinatura', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
            print('❌ Existe criticas nos dados que impedem a assinatura')
            
            excluir_documento()
            
            return False
            
        if _find_img('assinar_com_certificado.png', conf=0.9):
            p.hotkey('alt', 'n')
        if _find_img('processo_finalizado_ass.png', conf=0.9):
            p.press('enter')
        time.sleep(1)
    
    p.hotkey('alt', 't')
    
    # iniciar trasnmissão do arquivo
    _wait_img('transmissao.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.hotkey('alt', 'i')
    
    while not _find_img('identificacao.png', conf=0.9):
        time.sleep(1)
    _click_img('identificacao.png', conf=0.9)

    # usuário e senha do contribuinte para transmitir sedif sem movimento
    p.write('RPOSTAL')
    p.press('tab')
    p.write('f7j54kymq4')
    _click_img('ok.png', conf=0.9)
    
    while not _find_img('processo_finalizado.png', conf=0.9):
        time.sleep(1)
        if _find_img('CPFdoResponsavelinvalido.png', conf=0.9):
            p.hotkey('alt', 'f')
            time.sleep(1)
            p.hotkey('alt', 'f')
            time.sleep(1)
            p.press('f')
            
            excluir_documento()
            
            _escreve_relatorio_csv(f'{cnpj};{nome};CPF do responsável inválido', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
            print('❌ CPF do responsável inválido')
            return False
        
    p.press('enter')
    
    _wait_img('salvar_recibo.png', conf=0.9, timeout=-1)
    time.sleep(1)

    _click_img('salvar_recibo.png', conf=0.9)

    _wait_img('salvar_como.png', conf=0.9, timeout=-1)
    time.sleep(1)

    copy(f'Recibo  de entrega SN - {mes}-{ano} - {cnpj}')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    copy('V:\Setor Robô\Scripts Python\Geral\Gerar Cadastro DSTDA\execucao\Recibos')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    
    _wait_img('deseja_visualizar.png', conf=0.9, timeout=-1)
    _click_img('deseja_visualizar.png', conf=0.9)
    time.sleep(1)
    
    p.press('right')
    p.press('enter')

    _wait_img('recibo_gerado.png', conf=0.9, timeout=-1)
    time.sleep(1)

    p.press('enter')

    _wait_img('transmissao.png', conf=0.9, timeout=-1)
    _click_img('transmissao.png', conf=0.9)
    time.sleep(2)

    p.hotkey('alt', 'f')
    time.sleep(1)
    p.hotkey('alt', 'f')
    time.sleep(1)
    p.hotkey('alt', 'f')
    time.sleep(1)

    _wait_img('tela_inicial.png', conf=0.9, timeout=-1)

    _wait_img('iniciar.png', conf=0.9, timeout=-1)
    _click_img('iniciar.png', conf=0.9)
    p.hotkey('alt', 'f')
    time.sleep(1)
    p.press('f')
    
    excluir_documento()
    
    _escreve_relatorio_csv(f'{cnpj};{nome};Competência transmitida', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
    print('✔ Competência transmitida')
    
    return True


@_time_execution
def run():
    comp = p.prompt(title='Script incrível', text='Qual competência', default='00/0000')
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        reinstala_sedif()
        if not configura(empresa, comp):
            continue
        transmitir(empresa, comp)
        p.hotkey('alt', 'f4')
        time.sleep(1)
        p.hotkey('alt', 's')


if __name__ == '__main__':
    run()
