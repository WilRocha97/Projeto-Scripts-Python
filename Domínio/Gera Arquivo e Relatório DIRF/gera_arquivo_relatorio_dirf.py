# -*- coding: utf-8 -*-
import datetime, pyperclip, time, os, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login, _salvar_pdf, _login_web, _abrir_modulo


def dirf(processos, empresa, ano, andamento):
    cod, cnpj, nome = empresa
    
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Impostos
    p.press('i')
    time.sleep(0.5)
    p.press('enter')
    # federais
    time.sleep(0.5)
    p.press('a')
    # DIRF
    time.sleep(0.5)
    p.press('d')
    # a partir de 2010
    time.sleep(0.5)
    p.press('enter')
    
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

    time.sleep(0.5)
    if not p.locateOnScreen(r'imgs/escrita_selecionado.png'):
        _click_img('escrita.png', conf=0.99)

    # abre a janela de outros dados
    p.hotkey('alt', 'u')

    while not _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)
        if _find_img('dirf_gerada.png', conf=0.9):
            _click_img('dirf_gerada.png', conf=0.8)
            p.press('enter')
            time.sleep(2)
            # abre a janela de outros dados
            p.hotkey('alt', 'u')

    if _find_img('folha_de_pagamento.png', conf=0.95):
        p.press('esc')
        time.sleep(1)
        p.hotkey('alt', 'n')
        time.sleep(1)
        p.press('esc')

        _escreve_relatorio_csv(f'{cod};{cnpj};{nome};Não é possível editar aba de folha de pagamento', nome=andamento)
        print('❌ Não é possível editar aba de folha de pagamento')
        return False

    _click_img('tem_folha.png', conf=0.95)
    time.sleep(1)

    # clicar para gerar de todos os colaboradores
    if _find_img('todos.png', conf=0.95):
        _click_img('todos.png', conf=0.95)
    time.sleep(0.2)

    if _find_img('gerar_info_complementar.png', conf=0.99):
        _click_img('gerar_info_complementar.png', conf=0.99)
    time.sleep(0.2)

    if _find_img('limitar_60_caractere.png', conf=0.99):
        _click_img('limitar_60_caractere.png', conf=0.99)
    time.sleep(0.2)

    p.hotkey('alt', 'g')

    while not _find_img('dirf.png', conf=0.9):
        time.sleep(1)
    
    if processos == 'Arquivos':
        resultados = arquivos_dirf(cod, ano)
    elif processos == 'Relatórios':
        resultados = relatorio_dirf(cod, cnpj, nome, ano)
    else:
        resultado_arquivo = arquivos_dirf(cod, ano)
        resultado_relatorio = relatorio_dirf(cod, cnpj, nome, ano)
        resultados = f'{resultado_arquivo};{resultado_relatorio}'

    _escreve_relatorio_csv(f'{cod};{cnpj};{nome};{resultados}', nome=andamento)
    p.press('esc', presses=4)
    time.sleep(2)

    
def relatorio_dirf(cod, cnpj, nome, ano):
    print('>>> Gerando relatório')
    # seleciona para gerar o relatório
    if _find_img('dirf_gerada.png', conf=0.9):
        _click_img('dirf_gerada.png', conf=0.8)
        p.press('enter')
    
    if _find_img('formulario.png', conf=0.9):
        _click_img('formulario.png', conf=0.9)
    time.sleep(1)

    p.hotkey('alt', 'o')
    
    while True:
        if _find_img('relatorio_gerado.png', conf=0.9):
            break
        if _find_img('relatorio_gerado_2.png', conf=0.9):
            break
        if _find_img('sem_dados_arquivo.png', conf=0.9):
            p.press('enter')
            print('❗ Sem dados para emitir')
            return 'Sem dados para emitir'
    
    p.hotkey('ctrl', 'd')
    
    print('>>> Salvando relatório')
    cnpj = cnpj.replace('/', '').replace('.', '').replace('-', '')
    nome = nome.replace('/', ' ').replace('?', ' ').replace(':', ' ').replace('"', ' ').replace('*', ' ')
    while True:
        if _find_img('salvar_em_pdf.png', pasta='imgs_c', conf=0.9):
            _salvar_pdf(abriu_janela=True)
            mover_relatorio_2(cod, cnpj, nome)
            break
        if _find_img('procurar_pasta.png', conf=0.9):
            salvar_pdf()
            mover_relatorio(cod, cnpj, nome)
            break
        
    print('✔ Relatório gerado')
    return f'Relatório DIRF {ano} gerado'


def arquivos_dirf(cod, ano):
    caminho = 'M:\\'

    print('>>> Gerando arquivo')
    erro = ''

    # seleciona para gerar o relatório
    if _find_img('arquivo.png', conf=0.9):
        _click_img('arquivo.png', conf=0.9)
    time.sleep(0.5)

    while not _find_img('pasta_de_destino_correto.png', conf=0.9):
        # desce para a linha onde sera digitado o caminho para salvar o arquivo
        _click_position_img('pasta_de_destino.png', '+', pixels_x=70, conf=0.9, clicks=2)
        time.sleep(1)

        # apaga qualquer texto que esteja no campo
        p.press('del', presses=85)
        p.press('backspace', presses=85)

        # digita o caminho para salvar o arquivo
        while True:
            try:
                pyperclip.copy(caminho)
                pyperclip.copy(caminho)
                p.hotkey('ctrl', 'v')
                time.sleep(0.2)
                _click_img('arquivo.png', conf=0.9)
                break
            except:
                pass

    time.sleep(1)

    p.hotkey('alt', 'o')
    
    while not _find_img('dirf_gerada.png', conf=0.9):
        time.sleep(1)
        if _find_img('rodar_processo.png', conf=0.9):
            p.hotkey('enter')
            while not _find_img('movimento_de_notas.png', conf=0.9):
                time.sleep(1)
            p.hotkey('alt', 'g')
            while not _find_img('dados_gravados.png', conf=0.9):
                time.sleep(1)
            p.hotkey('enter')
        
        if _find_img('outros_dados_nao_digitados.png', conf=0.9):
            erro += ' - Outros dados não digitados'
            p.hotkey('enter')
        
        if _find_img('nao_tem_parametro.png', conf=0.9):
            erro += ' - Não existe parâmetro informado'
            p.hotkey('enter')
        
        if _find_img('erro_arquivo.png', conf=0.9):
            erro += ' - Não é possível gerar o arquivo'
            p.hotkey('enter')
            
        if _find_img('sem_dados_arquivo.png', conf=0.9):
            p.press('enter')
            print(f'❗ Sem dados para emitir{erro}')
            return f'Sem dados para emitir{erro}'

        if _find_img('responsavel_invalido.png', conf=0.9):
            p.press('enter')
            print(f'❗ Responsável inválido{erro}')
            return f'Responsável inválido{erro}'

    time.sleep(3)

    erro += mover_arquivo(cod)
    print(f'✔ Arquivo gerado{erro}')
    return f'Arquivo DIRF {ano} gerado{erro}'


def salvar_pdf():
    timer = 0
    while not _find_img('procurar_pasta.png', conf=0.9):
        time.sleep(1)
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
        os.makedirs('C:/Arquivos', exist_ok=True)
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

    return True


def mover_arquivo(cod):
    os.makedirs('execução/Arquivos', exist_ok=True)
    final_folder = "execução\\Arquivos"
    folder = 'C:\\'
    
    for arq in [f'DIRF{cod}.txt', f'DIRF0{cod}.txt', f'DIRF00{cod}.txt', f'DIRF000{cod}.txt', f'DIRF0000{cod}.txt', f'DIRF00000{cod}.txt']:
        try:
            shutil.move(os.path.join(folder, arq), os.path.join(final_folder, f'DIRF_{cod}.txt'))
            return ''
        except:
            pass
    
    return ' - Erro ao mover o arquivo'


def mover_relatorio(cod, cnpj, nome):
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


def mover_relatorio_2(cod, cnpj, nome):
    os.makedirs('execução/Relatórios', exist_ok=True)

    folder = "V:\\Setor Robô\\Scripts Python\\Domínio\\Gera Arquivo e Relatório DIRF\\execução\\Relatórios"
    guia_folha = os.path.join('C:\\DIRF Folha.pdf')
    guia_escrita = os.path.join('C:\\DIRF Escrita.pdf')
    while True:
        if os.path.exists(guia_folha):
            guia = guia_folha
            break
        if os.path.exists(guia_escrita):
            guia = guia_escrita
            break
        time.sleep(1)

    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(folder, f'DIRF{cod} - {cnpj} - {nome}.pdf'))
            time.sleep(2)
        except:
            pass


@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('folha')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
    
        if not _login(empresa, andamentos):
            continue
        dirf(processos, empresa, ano, andamentos)


if __name__ == '__main__':
    ano = p.prompt(text='Qual ano base?', title='Script incrível', default='0000')
    processos = p.confirm(buttons=['Arquivos', 'Relatórios', 'Arquivos e Relatórios'])
    empresas = _open_lista_dados()
    andamentos = 'Arquivos DIRF'

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
