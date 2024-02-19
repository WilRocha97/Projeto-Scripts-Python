# -*- coding: utf-8 -*-
import fitz, re, shutil, time, os, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _ask_for_dir, _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf


def relatorio_darf_dctf(empresa, periodo, andamento):
    cod, cnpj, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('i')
    # Resumo
    time.sleep(0.5)
    p.press('m')
    verificacao = 'continue'
    while not _find_img('resumo_de_impostos.png', conf=0.9):
        
        if verificacao != 'continue':
            return verificacao, ''
        time.sleep(1)
        if _find_img('vigencia_sem_parametro.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro']), nome=andamento)
            print('❌ Não existe parametro')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok', ''

    # Período
    p.write(periodo)
    p.press('tab')
    time.sleep(1)
    
    if _find_img('sem_parametro_vigencia.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro para a vigência: ' + periodo]), nome=andamento)
        print('❌ Não existe parametro para a vigência: ' + periodo)
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        p.press('esc')
        time.sleep(1)
        return 'ok', ''
    
    p.write(periodo)
    time.sleep(0.5)

    # Todos os impostos
    p.hotkey('alt', 't')
    time.sleep(1)

    while _find_img('destacar_linhas.png', conf=0.95):
        _click_img('destacar_linhas.png', conf=0.95)
        time.sleep(0.5)
        
    '''while _find_img('detalhar_dados.png', conf=0.95):
                    _click_img('detalhar_dados.png', conf=0.95)
                    time.sleep(0.5)'''

    p.hotkey('alt', 'o')
    time.sleep(1)
    sem_layout = 0

    while not _find_img('resumo_gerado.png', conf=0.8):
        if verificacao != 'continue':
            return verificacao, ''
        time.sleep(1)
        if _find_img('imposto_sem_layout.png', conf=0.9):
            sem_layout = 1
            p.press('enter')
        if sem_layout == 1:
            while not _find_img('resumo_de_impostos.png', conf=0.9):
                if verificacao != 'continue':
                    return verificacao, ''
                p.press('enter')
                time.sleep(1)
            p.press('esc', presses=4)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Imposto sem layout']), nome=andamento)
            print('❌ Imposto sem layout')
            return 'ok', ''

        time.sleep(3)
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok', ''

        if _find_img('sem_imposto.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return 'ok', ''

        if _find_img('resumo_calculado.png', conf=0.9) or _find_img('resumo_gerado_2.png', conf=0.9):
            break

    _salvar_pdf()
    arq_final = mover_relatorio(cod)
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório gerado']), nome=andamento)
    print('✔ Relatório gerado')

    p.press('esc', presses=4)
    time.sleep(2)
    
    return 'ok', arq_final


def mover_relatorio(cod):
    folder = 'C:\\'
    for arq in os.listdir(folder):
        if arq.endswith('.pdf'):
            print(arq)
            if re.compile(r'Empresa ' + str(cod)).search(arq):
                os.makedirs('execução/Relatórios', exist_ok=True)
                final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Relatórios para DARF DCTF\\execução\\Relatórios'
                arq_final = os.path.join(final_folder, arq)
                shutil.move(os.path.join(folder, arq), arq_final)
                # analisa_arquivo(cod, nome, arq_final)


def analisa_arquivo(cod, nome, arq_final):
    doc = fitz.open(arq_final, filetype="pdf")
    texto = ''
    total_competencia_lancados = False
    total_competencia_calculados = False
    # print(arq_final)
    for page in doc:
        
        texto_pagina = page.get_text('text', flags=1 + 2 + 8)
        texto += texto_pagina
        
    texto = texto.replace('1Total competência', 'Total competência')
    cnpj = re.compile(r'\d\d/\d\d\d\d\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)').search(texto)
    """if re.compile('IESMAR SARAIVA 04541756803').search(texto):
        print(texto)
        time.sleep(22)"""
    if cnpj:
        cnpj = cnpj.group(1)
        
        for i in range(1000):
            total_lancados = re.compile(r'RESUMO DOS IMPOSTOS LANÇADOS(\n.+){' + str(i) + '}\n.+otal geral:\n(.+)').search(texto)
            if total_lancados:
                total_competencia_lancados = total_lancados.group(2)
                break
        
        for i in range(1000):
            total_calculados = re.compile(r'RESUMO DOS IMPOSTOS CALCULADOS(\n.+){' + str(i) + '}\n.+otal geral:\n(.+)').search(texto)
            if total_calculados:
                total_competencia_calculados = total_calculados.group(2)
                break

        if total_competencia_lancados and not total_competencia_calculados:
            # print('Lançados')
            impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto)
            for imposto in impostos:
                imposto = imposto[9].replace('(', '\(').replace(')', '\)')
                linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
                if linha:
                    imposto_1 = linha.group(1)
                    imposto_2 = linha.group(2)
                    imposto_3 = linha.group(3)
                    imposto_4 = linha.group(4)
                    imposto_5 = linha.group(6)
                    imposto_6 = linha.group(7)
                    imposto_7 = linha.group(8)
                    imposto_8 = linha.group(9)
                    imposto_recolher = linha.group(5)
                    imposto_nome = linha.group(10)
                    linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS LANÇADOS;{imposto_nome};{imposto_1};{imposto_2};{imposto_3};{imposto_4};{imposto_5};{imposto_6};{imposto_7};{imposto_8};{imposto_recolher}'
                    _escreve_relatorio_csv(f'{linha_};{total_competencia_lancados}', nome='Resumo relatórios')
                  
        elif total_competencia_calculados and not total_competencia_lancados:
            # print('Calculados')
            impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto)
            for imposto in impostos:
                imposto = imposto[3].replace('(', '\(').replace(')', '\)')
                if impostos:
                    linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
                    if linha:
                        imposto_recolher = linha.group(2)
                        imposto_nome = linha.group(4)
                        linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS CALCULADOS;{imposto_nome};{imposto_recolher}'
                        _escreve_relatorio_csv(f'{linha_};{total_competencia_calculados}', nome='Resumo relatórios')
                    
                    linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
                    if linha:
                        imposto_1 = linha.group(1)
                        imposto_2 = linha.group(2)
                        imposto_3 = linha.group(3)
                        imposto_4 = linha.group(4)
                        imposto_5 = linha.group(5)
                        imposto_6 = linha.group(7)
                        imposto_7 = linha.group(8)
                        imposto_recolher = linha.group(6)
                        imposto_nome = linha.group(9)
                        linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS CALCULADOS;{imposto_nome};{imposto_1};{imposto_2};{imposto_3};{imposto_4};{imposto_5};{imposto_6};{imposto_7};{imposto_recolher}'
                        _escreve_relatorio_csv(f'{linha_};{total_competencia_calculados}', nome='Resumo relatórios')
                        
                    linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
                    if linha:
                        imposto_recolher = linha.group(6)
                        imposto_nome = linha.group(11)
                        linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS CALCULADOS;{imposto_nome};{imposto_recolher}'
                        _escreve_relatorio_csv(f'{linha_};{total_competencia_calculados}', nome='Resumo relatórios')
  
        elif total_competencia_lancados and total_competencia_calculados:
            # print('Lançados e calculados')
            texto_dividido = texto.split('RESUMO DOS IMPOSTOS CALCULADOS')
            
            impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto_dividido[0])
            for imposto in impostos:
                imposto = imposto[9].replace('(', '\(').replace(')', '\)')
                linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto_dividido[0])
                if linha:
                    imposto_1 = linha.group(1)
                    imposto_2 = linha.group(2)
                    imposto_3 = linha.group(3)
                    imposto_4 = linha.group(4)
                    imposto_5 = linha.group(6)
                    imposto_6 = linha.group(7)
                    imposto_7 = linha.group(8)
                    imposto_8 = linha.group(9)
                    imposto_recolher = linha.group(5)
                    imposto_nome = linha.group(10)
                    linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS LANÇADOS;{imposto_nome};{imposto_1};{imposto_2};{imposto_3};{imposto_4};{imposto_5};{imposto_6};{imposto_7};{imposto_8};{imposto_recolher}'
                    _escreve_relatorio_csv(f'{linha_};{total_competencia_lancados}', nome='Resumo relatórios')
            
            impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto_dividido[1])
            for imposto in impostos:
                imposto = imposto[3].replace('(', '\(').replace(')', '\)')
                if impostos:
                    linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto_dividido[1])
                    if linha:
                        imposto_1 = linha.group(1)
                        imposto_2 = linha.group(3)
                        imposto_recolher = linha.group(2)
                        imposto_nome = linha.group(4)
                        linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS CALCULADOS;{imposto_nome};{imposto_1};{imposto_2};{imposto_recolher}'
                        _escreve_relatorio_csv(f'{linha_};{total_competencia_calculados}', nome='Resumo relatórios')

                    linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto_dividido[1])
                    if linha:
                        imposto_1 = linha.group(1)
                        imposto_2 = linha.group(2)
                        imposto_3 = linha.group(3)
                        imposto_4 = linha.group(4)
                        imposto_5 = linha.group(5)
                        imposto_6 = linha.group(7)
                        imposto_7 = linha.group(8)
                        imposto_8 = linha.group(9)
                        imposto_9 = linha.group(10)
                        imposto_recolher = linha.group(6)
                        imposto_nome = linha.group(11)
                        linha_ = f'{cod};{cnpj};{nome};RESUMO DOS IMPOSTOS CALCULADOS;{imposto_nome};{imposto_1};{imposto_2};{imposto_3};{imposto_4};{imposto_5};{imposto_6};{imposto_7};{imposto_8};{imposto_9};{imposto_recolher}'
                        _escreve_relatorio_csv(f'{linha_};{total_competencia_calculados}', nome='Resumo relatórios')

        else:
            print(texto)
            p.alert(text='Erro')
            
    print('>>> Arquivo analisado')
    return True


@_time_execution
# @_barra_de_status
def run():
    if rotina == 'Sim':
        for arq in os.listdir(pasta):
            arquivo = os.path.join(pasta, arq)
            arq_name = re.compile(r'Empresa (\d+) - (.+).pdf').search(arq)
            cod = arq_name.group(1)
            nome = arq_name.group(2)
            
            analisa_arquivo(cod, nome, arquivo)
            
    else:
        _login_web()
        _abrir_modulo('escrita_fiscal')
        
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            # printa o indice da empresa que está sendo executada
            _indice(count, total_empresas, empresa, index)
    
            while True:
                if not _login(empresa, andamentos):
                    break
                else:
                    resultado, arq_final = relatorio_darf_dctf(empresa, periodo, andamentos)
                    
                    if resultado == 'dominio fechou':
                        _login_web()
                        _abrir_modulo('escrita_fiscal')
                
                    if resultado == 'modulo fechou':
                        _abrir_modulo('escrita_fiscal')
                    
                    if resultado == 'ok':
                        break


if __name__ == '__main__':
    rotina = p.confirm(text='Gerar resumo dos relatórios já salvos?', buttons=('Sim', 'Não'))
    if rotina == 'Sim':
        pasta = _ask_for_dir()
        run()
    else:
        periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/0000')
        empresas = _open_lista_dados()
        andamentos = 'Relatórios Resumo de Impostos'
    
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is not None:
            run()
            