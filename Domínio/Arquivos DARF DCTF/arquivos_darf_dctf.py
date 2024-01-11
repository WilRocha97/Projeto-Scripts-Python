# -*- coding: utf-8 -*-
import os
import time, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login


def arquivos_darf_dctf(empresa, periodo, andamento):
    cod, cnpj,  nome = empresa
    nome_arquivo = 'M:\DCTF_{}.RFB'.format(cod)
    
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios mensal
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('n')
    # Resumo
    time.sleep(0.5)
    p.press('f')
    time.sleep(0.5)
    p.press('d')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('m')

    while not _find_img('dctf_mensal.png', conf=0.9):
        if _find_img('dctf_mensal_2.png', conf=0.9):
            break
        time.sleep(1)

    # Período
    p.write(periodo)
    time.sleep(1)

    p.press('tab')
    time.sleep(1)

    p.press('delete', presses=25)
    time.sleep(1)

    p.write(nome_arquivo)
    time.sleep(1)

    p.hotkey('alt', 'o')

    while not _find_img('outros_dados.png', conf=0.9):
        if _find_img('outros_dados_2.png', conf=0.9):
            break
            
        if _find_img('data_inicio.png', conf=0.9):
            _click_img('data_inicio.png', conf=0.9)
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Competência informada menor que a data de início efetivo das atividades.']), nome=andamento)
            print('❗ Competência informada menor que a data de início efetivo das atividades.')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo
        
        time.sleep(2)

    p.click(1214, 488)
    time.sleep(2)

    p.write('PJ nao')
    time.sleep(2)
    p.press('enter')
    time.sleep(2)
    p.click(1214, 488)
    time.sleep(2)

    p.click(413, 541)
    time.sleep(2)

    p.click(1215, 547)
    time.sleep(2)
    if _find_img('sem_alteracao.png', conf=0.99):
        _click_img('sem_alteracao.png', conf=0.99)
        time.sleep(1)

    _click_img('ok.png', conf=0.9)
    time.sleep(1)

    while _find_img('outros_dados.png', conf=0.9):
        time.sleep(2)

    p.hotkey('alt', 'x')
    print('>>> Gerando arquivo')
    while _find_img('dctf_mensal.png', conf=0.9) or _find_img('dctf_mensal_2.png', conf=0.9):
        time.sleep(1)
        if _find_img('nao_gerou_arquivo.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não gerou arquivo']), nome=andamento)
            print('❌ Não gerou arquivo')
            time.sleep(1)
            p.press('esc', presses=5)

        if _find_img('nao_gerou_arquivo_2.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não gerou arquivo']), nome=andamento)
            print('❌ Não gerou arquivo')
            time.sleep(1)
            p.press('esc', presses=5)
        
        if _find_img('imune_irpj.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Imune do IRPJ']), nome=andamento)
            print('❌ Imune do IRPJ')
            time.sleep(1)

        if _find_img('isenta.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Isenta do IRPJ']), nome=andamento)
            print('❌ Isenta do IRPJ')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo

        if _find_img('saldo_nao_calculado.png', conf=0.9) or _find_img('saldo_nao_calculado_2.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Saldo dos impostos não foi calculado no período']), nome=andamento)
            print('❌ Saldo dos impostos não foi calculado no período')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo

        if _find_img('nao_tem_parametro.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro para a vigência ', periodo]), nome=andamento)
            print('❌ Não existe parametro para a vigência {}'.format(periodo))
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo

        if _find_img('exportacao_cancelada.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo

        if _find_img('final_da_exportacao.png', conf=0.9) or _find_img('final_da_exportacao_2.png', conf=0.9):
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Arquivo gerado']), nome=andamento)
            print('✔ Arquivo gerado')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'arquivo gerado', nome_arquivo
        
    print('❗ Erro inesperado')
    p.press('esc', presses=5)
    time.sleep(3)
    return 'ok', nome_arquivo


def mover_arquivo(nome_arquivo):
    nome_arquivo = nome_arquivo.replace('M:\DCTF', 'DCTF')
    os.makedirs('execução/Arquivos', exist_ok=True)
    final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Arquivos DARF DCTF\\execução\\Arquivos'
    folder = 'C:\\'
    shutil.move(os.path.join(folder, nome_arquivo), os.path.join(final_folder, nome_arquivo))


@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count + index)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index))} Restantes')
        _indice(count, total_empresas, empresa, index)
        
        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado, nome_arquivo = arquivos_darf_dctf(empresa, periodo, andamentos)
                if resultado == 'arquivo gerado':
                    mover_arquivo(nome_arquivo)
                    break
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
    
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break
    

if __name__ == '__main__':
    periodo = p.prompt(text='Qual o período do arquivo', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos para DARF DCTF'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        