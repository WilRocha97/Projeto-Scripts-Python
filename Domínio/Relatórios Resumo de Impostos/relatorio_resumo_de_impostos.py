# -*- coding: utf-8 -*-
import datetime, fitz, re, shutil, time, os, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _ask_for_dir, _indice, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
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
    arq_final = mover_relatorio(cod, nome, periodo)
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório gerado']), nome=andamento)
    print('✔ Relatório gerado')

    p.press('esc', presses=4)
    time.sleep(2)
    
    return 'ok', arq_final


def mover_relatorio(cod, nome, periodo):
    folder = 'C:\\'
    for arq in os.listdir(folder):
        if arq.endswith('.pdf'):
            print(arq)
            if re.compile(r'Empresa ' + str(cod)).search(arq):
                os.makedirs(f'execução/Relatórios {periodo.replace("/", "-")}', exist_ok=True)
                final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Relatórios Resumo de Impostos\\execução\\Relatórios ' + periodo.replace('/', '-')
                arq_final = os.path.join(final_folder, arq)
                shutil.move(os.path.join(folder, arq), arq_final)
                analisa_arquivo(cod, nome, periodo, arq_final)


def analisa_arquivo(cod, nome, periodo, arq_final):
    doc = fitz.open(arq_final, filetype="pdf")
    texto = ''
    total_competencia_lancados = False
    total_competencia_calculados = False
    # print(arq_final)
    # captura o texto de todas as páginas e salva em uma única variável
    for page in doc:
        texto_pagina = page.get_text('text', flags=1 + 2 + 8)
        texto += texto_pagina
    
    # captura o cnpj da empresa
    cnpj = re.compile(r'\d\d/\d\d\d\d\n(\d\d.\d\d\d.\d\d\d/\d\d\d\d-\d\d)').search(texto)

    if cnpj:
        cnpj = cnpj.group(1)
        
        # laço para capturar o valor total dos impostos lançados
        for i in range(1000):
            total_lancados = re.compile(r'RESUMO DOS IMPOSTOS LANÇADOS(\n.+){' + str(i) + '}\n.+otal geral:\n(.+)').search(texto)
            if total_lancados:
                total_competencia_lancados = total_lancados.group(2)
                break
        
        # laço para capturar o valor total dos impostos calculados
        for i in range(1000):
            total_calculados = re.compile(r'RESUMO DOS IMPOSTOS CALCULADOS(\n.+){' + str(i) + '}\n.+otal geral:\n(.+)').search(texto)
            if total_calculados:
                total_competencia_calculados = total_calculados.group(2)
                break
        
        # se encontrar o total dos lançados e NÃO dos calculados
        if total_competencia_lancados and not total_competencia_calculados:
            captura_imposto_lacados(cnpj, cod, nome, periodo, texto, total_competencia_lancados)
        
        # se encontrar o total dos calculados e NÂO dos lançados
        elif total_competencia_calculados and not total_competencia_lancados:
            captura_imposto_calculados(cnpj, cod, nome, periodo, texto, total_competencia_calculados)
    
        # se encontrar os dois
        elif total_competencia_lancados and total_competencia_calculados:
            # divide o texto em duas partes para que cada bloco de lançados e calculados seja analisado individualmente
            texto_dividido = texto.split('RESUMO DOS IMPOSTOS CALCULADOS')
            
            captura_imposto_lacados(cnpj, cod, nome, periodo, texto_dividido[0], total_competencia_lancados)
            captura_imposto_calculados(cnpj, cod, nome, periodo, texto_dividido[1], total_competencia_calculados)

        else:
            print(texto)
            p.alert(text='Erro')
            
    print('>>> Arquivo analisado')
    return True


def captura_imposto_lacados(cnpj, cod, nome, periodo, texto, total_competencia):
    # captura a lista de impostos
    impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto)
    # para cada imposto realiza o processo
    for imposto in impostos:
        # formata o texto do regex para caso o nome do imposto contenha "(" ou ")" ele entenda que é um caractere e não agrupamento
        imposto = imposto[9].replace('(', '\(').replace(')', '\)')
        linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
        # se encontrar a linha do imposto monta a linha para anotar na planilha
        if linha:
            insere_linha(cod, cnpj, nome, periodo, total_competencia, resumo_nome='RESUMO DOS IMPOSTOS LANÇADOS', debitos=linha.group(1), creditos=linha.group(2), acrescimos=linha.group(3),
                         outras_deducoes=linha.group(4), imposto_recolher=linha.group(5), imposto_diferido=linha.group(6), saldo_credor=linha.group(7),
                         saldo_credor_anterior=linha.group(8), saldo_diferido_anterior=linha.group(9), imposto_nome=linha.group(10))


def captura_imposto_calculados(cnpj, cod, nome, periodo, texto, total_competencia):
    # captura a lista de impostos
    impostos = re.compile(r'(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n([A-Z]+.+)').findall(texto)
    # para cada imposto realiza o processo
    for imposto in impostos:
        # formata o texto do regex para caso o nome do imposto contenha "(" ou ")" ele entenda que é um caractere e não agrupamento
        imposto = imposto[3].replace('(', '\(').replace(')', '\)')
        linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
        # se encontrar a linha do imposto monta a linha para anotar na planilha com algumas variações de quantidade de colunas no PDF
        if linha:
            insere_linha(cod, cnpj, nome, periodo, total_competencia, resumo_nome='RESUMO DOS IMPOSTOS CALCULADOS', imposto_diferido=linha.group(1), imposto_recolher=linha.group(2),
                         saldo_credor=linha.group(3), imposto_nome=linha.group(4))
        
        linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
        if linha:
            insere_linha(cod, cnpj, nome, periodo, total_competencia, resumo_nome='RESUMO DOS IMPOSTOS CALCULADOS', imposto_recolher=linha.group(6), imposto_nome=linha.group(9))
        
        linha = re.compile(r'^[^,]*$\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(.+,\d+)\n(' + str(imposto) + ')', re.MULTILINE).search(texto)
        if linha:
            insere_linha(cod, cnpj, nome, periodo, total_competencia, resumo_nome='RESUMO DOS IMPOSTOS CALCULADOS', acrescimos=linha.group(1), outras_deducoes=linha.group(2), saldo_credor_anterior=linha.group(3),
                         saldo_diferido_anterior=linha.group(4), imposto_diferido=linha.group(5), imposto_recolher=linha.group(6), valor_imposto=linha.group(7),
                         base_calculo=linha.group(8), aliquota=linha.group(9), saldo_credor=linha.group(10), imposto_nome=linha.group(11))


def insere_linha(cod, cnpj, nome, periodo, total_competencia, resumo_nome='resumo', imposto_nome='nome', base_calculo='0', aliquota='0', valor_imposto='0', saldo_credor_anterior='0', saldo_diferido_anterior='0',
                 debitos='0', creditos='0', acrescimos='0', outras_deducoes='0', imposto_recolher='0', imposto_diferido='0', saldo_credor='0'):
    linha_ = (f'{cod};{cnpj};{nome};{resumo_nome};{imposto_nome};'
              f'{base_calculo};'
              f'{aliquota};'
              f'{valor_imposto};'
              f'{saldo_credor_anterior};'
              f'{saldo_diferido_anterior};'
              f'{debitos};'
              f'{creditos};'
              f'{acrescimos};'
              f'{outras_deducoes};'
              f'{imposto_recolher};'
              f'{imposto_diferido};'
              f'{saldo_credor}')
    _escreve_relatorio_csv(f'{linha_};{total_competencia}', nome=f'Resumo relatórios {periodo.replace("/", "-")}')


@_time_execution
@_barra_de_status
def run(window):
    if rotina == 'Sim':
        for arq in os.listdir(pasta):
            arquivo = os.path.join(pasta, arq)
            arq_name = re.compile(r'Empresa (\d+) - (.+).pdf').search(arq)
            cod = arq_name.group(1)
            nome = arq_name.group(2)
            
            analisa_arquivo(cod, nome, arquivo, window)

    else:
        _login_web()
        _abrir_modulo('escrita_fiscal')
        
        tempos = [datetime.datetime.now()]
        tempo_execucao = []
        total_empresas = empresas[index:]
        for count, empresa in enumerate(empresas[index:], start=1):
            # printa o indice da empresa que está sendo executada
            tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
    
            while True:
                if not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
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
    
    _escreve_header_csv('COD;CNPJ;NOME;RESUMO;IMPOSTO;BASE CALCULO;ALIQUOTA;VALOR IMPOSTO;SALDO CREDOR ANTERIOR;SALDO DEFERIDO ANTERIOR;DÉBITOS;CRÉDITOS;ACRESCIMOS;'
                        'OUTRAS DEDUÇÕES;IMPOSTO RECOLHER;IMPOSTO DEFERIDO;SALDO CREDOR;TOTAL', nome=f'Resumo relatórios {periodo.replace("/", "-")}')
    

if __name__ == '__main__':
    rotina = p.confirm(text='Gerar resumo dos relatórios já salvos?', buttons=('Sim', 'Não'))
    if rotina == 'Sim':
        pasta = _ask_for_dir()
        run()
    else:
        periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='00/0000')
        empresas = _open_lista_dados()
        andamentos = f'Relatórios Resumo de Impostos {periodo.replace("/", "-")}'
    
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is not None:
            run()
            