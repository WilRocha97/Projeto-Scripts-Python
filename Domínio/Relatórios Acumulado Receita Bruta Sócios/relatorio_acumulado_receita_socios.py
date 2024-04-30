# -*- coding: utf-8 -*-
import datetime, fitz, re, shutil, time, os, pyautogui as p
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf


def relatorio_acumulado(empresa, periodo, andamento):
    cod, cnpj, nome = empresa
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Acompanhamentos
    p.press('a')
    # Resumo
    time.sleep(0.5)
    p.press('b')

    while not _find_img('receita_bruta_global.png', conf=0.9):
        time.sleep(1)
    
    p.write(periodo)
    time.sleep(0.5)
    p.hotkey('alt', 'o')
    
    time.sleep(1)
    if _find_img('nao_simples.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não é Simples Nacional, não pode gerar o relatório']), nome=andamento)
        print('❌ Empresa não é Simples Nacional, não pode gerar o relatório')
        p.press('enter')
        time.sleep(1)
        p.press('esc')
        time.sleep(1)
        return
    
    resultado = espera_gerar(empresa, andamento)
    if not resultado:
        return
    
    _salvar_pdf()
    time.sleep(1)
    
    final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Relatórios Acumulado Receita Bruta Sócios\\execução\\Relatórios'
    
    mover_relatorio(final_folder)
    analisa_relatorio(final_folder, cnpj, nome)
    
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório gerado']), nome=andamento)
    print('✔ Relatório gerado')
    
    p.press('esc', presses=4)
    time.sleep(2)
    
    return


def mover_relatorio(final_folder):
    print('>>> Movendo relatório...')
    os.makedirs('execução/Relatórios', exist_ok=True)
    folder = 'C:\\'
    arq = 'Simples Nacional - Receita Bruta Global Acumulada.pdf'
    shutil.move(os.path.join(folder, arq), os.path.join(final_folder, arq))
    

def analisa_relatorio(final_folder, cnpj, nome_consulta):
    print('>>> Analisando relatório...')
    arq = 'Simples Nacional - Receita Bruta Global Acumulada.pdf'
    planilha = []
    with (fitz.open(os.path.join(final_folder, arq)) as pdf):
        
        # Para cada página do pdf, se for a segunda página o script ignora
        for count, page in enumerate(pdf):
            if count == 1:
                continue
            # Pega o texto da pagina
            textinho = page.get_text('text', flags=1 + 2 + 8)
            
            #print(textinho)
            #time.sleep(33)
            
            socios = re.compile(r'Sócio:(.+)').findall(textinho)
            
            for socio in socios:
                for i in range(150):
                    empresa = re.compile(r'Sócio:' + socio + '(\n.+){' + str(i) + '}\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+,\d\d)\n(.+)\n(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d\-\d\d)').search(textinho)
                
                    if empresa:
                        cnpj_lista = empresa.group(16)
                        nome = empresa.group(15)
                        jan = empresa.group(14)
                        fev = empresa.group(2)
                        mar = empresa.group(3)
                        abr = empresa.group(4)
                        mai = empresa.group(12)
                        jun = empresa.group(11)
                        jul = empresa.group(10)
                        ago = empresa.group(9)
                        sep = empresa.group(8)
                        out = empresa.group(7)
                        nov = empresa.group(6)
                        dez = empresa.group(13)
                        total = empresa.group(5)
                        
                        linha = f'{cnpj};{nome_consulta};{socio};{cnpj_lista};{nome};{jan};{fev};{mar};{abr};{mai};{jun};{jul};{ago};{sep};{out};{nov};{dez};{total}'
                        
                        if linha not in planilha:
                            planilha.append(linha)
    
    new_arq = f'Base da consulta= {cnpj} - Receita bruta global.pdf'
    shutil.move(os.path.join(final_folder, arq), os.path.join(final_folder, new_arq))
    for linha in planilha:
        _escreve_relatorio_csv(linha)
    

def espera_gerar(empresa, andamento):
    cod, cnpj, nome = empresa
    timer = 0
    # espera gerar
    while not _find_img('receita_bruta_gerada.png', conf=0.9):
        print('>>> Aguardando gerar')

        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=4, interval=1)
            time.sleep(1)
            return False

        if _find_img('sem_parametro.png', conf=0.9):
            p.press('enter')
            time.sleep(1)

        time.sleep(1)
        timer += 1
        
        if timer >= 30:
            p.hotkey('alt', 'o')
            timer = 0
    
    return True


@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        if not _login(empresa, andamentos):
            continue

        relatorio_acumulado(empresa, periodo, andamentos)

    
if __name__ == '__main__':
    periodo = p.prompt(text='Qual o período do relatório', title='Script incrível', default='0000')
    empresas = _open_lista_dados()
    andamentos = 'Relatórios Acumulado Sócios'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
