# -*- coding: utf-8 -*-
from datetime import datetime
import pyautogui as p
import time, os, shutil

from sys import path
path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _inicial, _iniciar
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def gerar_dirf(index, empresas):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod, cnpj, nome, mes, ano = empresa
        
        # Verificar horário
        _hora_limite = datetime.now().replace(hour=18, minute=25, second=0, microsecond=0)
        if _horario(_hora_limite, 'DPCUCA'):
            _iniciar('DPCUCA')
            p.getWindowsWithTitle('DPCUCA')[0].maximize()
        
        _indice(count, total_empresas, empresa)
        
        _inicial('dpcuca')

        # Verificações de login
        if not _login(empresa, 'Codigo', 'dpcuca', 'Gerador de arquivos DIRF', mes, ano):
            p.press('enter')
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Empresa Inativa', nome='Gerador de arquivos DIRF')
            print('❌ Empresa Inativa')
            continue

        # CNPJ com os separadores para poder verificar a empresa no cuca
        if not _verificar_empresa(cnpj, 'dpcuca'):
            _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{mes};{ano};Empresa não encontrada no DPCUCA', nome='Gerador de arquivos DIRF')
            print('❌ Empresa não encontrada no DPCUCA')
            continue

        time.sleep(4)
        _click_img('CUCAInicial.png', conf=0.9)
        # Abrir Calculos
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press(['d', 'd', 'enter'], interval=0.5)
        _wait_img('GerarInfo.png', conf=0.9)
        _click_img('Pasta.png', conf=0.9)
        _wait_img('Pasta.png', conf=0.9)
        time.sleep(0.5)
        p.press(['tab', 'tab', 'tab'], interval=0.5)
        time.sleep(0.5)
        p.write('T:\DIRF 2023-2022')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
        if _find_img('GerarBenefciarios.png', conf=0.9):
            _click_img('GerarBenefciarios.png', conf=0.9)
        time.sleep(0.5)
        _click_img('Gerar.png', conf=0.9)

        p.hotkey('alt', 'a')
        time.sleep(1)
        if _find_img('Atencao.png', conf=0.9):
            p.press('enter')
        time.sleep(2)
        while not _find_img('Inicio.png', conf=0.9):
            time.sleep(1)
            p.moveTo(683, 384)
            p.moveTo(684, 383)
            if _find_img('OK.png', conf=0.9):
                p.press('enter')
                time.sleep(0.5)
            if _find_img('Situacoes.png', conf=0.9):
                p.press('enter')
                time.sleep(0.5)
            if _find_img('DIRFgerada.png', conf=0.9):
                p.press('enter')
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'Arquivo DIRF gerado!']), nome='Gerador de arquivos DIRF')
                print('✔ Arquivo DIRF gerado')
            if _find_img('NaoDIRF.png', conf=0.9):
                p.press('enter')
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, mes, ano, 'Não foram encontrados informações para DIRF!']), nome='Gerador de arquivos DIRF')
                print('❌ Não foram encontrados informações para DIRF')

        _click_img('JanelaAGUARDE.png', conf=0.9)
        p.press('esc')

        _inicial('dpcuca')

        '''download_folder = "T:\\DIRF 2023-2022\\DIRF"
        folder = "\\vpsrv03\\Arq_Robo\\Gera Arquivos DIRF\\DIRF 2022-2023\\Folha\\Arquivos"
        for file in os.listdir(download_folder):
            if file.endswith('.txt'):
                shutil.move(os.path.join(download_folder, file), os.path.join(folder, file))'''
            

@_time_execution
def run():
    # Perguntar qual relatório
    inicio = datetime.now()

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    try:
        gerar_dirf(index, empresas)
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


if __name__ == '__main__':
    run()
    