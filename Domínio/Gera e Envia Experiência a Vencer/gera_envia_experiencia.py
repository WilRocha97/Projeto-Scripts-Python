# -*- coding: utf-8 -*-
import shutil, fitz, re, pyperclip, time, os, pyautogui as p
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, e_dir, _where_to_start
from dominio_comum import _login


def escreve_dados(cod, nome):
    f = open(os.path.join('ignore', 'Dados.csv'), 'a', encoding='latin-1')
    f.write(f'{cod};CNPJ;{nome}\n')
    f.close()
    
    
def open_lista_dados(file, encode='latin-1'):
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def salva_pdf_geral():
    _click_img('salvar_pdf.png', conf=0.9)
    while not _find_img('salvar_em_pdf.png', conf=0.9):
        time.sleep(1)
    
    if not _find_img('cliente_c_selecionado.png', conf=0.9):
        while not _find_img('cliente_c.png', conf=0.9):
            _click_img('botao.png', conf=0.9)
            time.sleep(3)
        _click_img('cliente_c.png', conf=0.9)
        time.sleep(5)
    
    p.press('enter')
    
    timer = 0
    while not _find_img('pdf_aberto.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('adobe.png', conf=0.9):
            p.press('enter')
        time.sleep(1)
        timer += 1
        if timer > 30:
            _click_img('salvar_pdf.png', conf=0.9)
            while not _find_img('salvar_em_pdf.png', conf=0.9):
                time.sleep(1)
            
            if not _find_img('cliente_c_selecionado.png', conf=0.9):
                while not _find_img('cliente_c.png', conf=0.9):
                    _click_img('botao.png', conf=0.9)
                    time.sleep(3)
                _click_img('cliente_c.png', conf=0.9)
                time.sleep(5)
            
            p.press('enter')
            timer = 0
    
    p.hotkey('alt', 'f4')
    time.sleep(2)

    shutil.move('C:\Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf', os.path.join('ignore', 'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf'))
    print('✔ Relatório gerado')
    return True


def gera_arquivo(andamento, cod='*', cnpj='', nome=''):
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        p.press('i', presses=2)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(2)
    
    time.sleep(0.5)
    p.write('Relação de Empregados')
    time.sleep(0.5)
    p.press('enter')
    _wait_img('relatorio_modelo_veiga.png', conf=0.9, timeout=-1)
    _click_img('relatorio_modelo_veiga.png', conf=0.9)
    
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(cod)

    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.press('*')

    time.sleep(0.5)
    p.hotkey('alt', 'e')
    
    while not _find_img('contrato_experiencia.png', conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para emitir']), nome=andamento)
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False
    
    if cod == '*':
        salva_pdf_geral()
    else:
        if not _find_img('enviar_arquivo.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não possuí opção de enviar o relatório para o cliente']), nome=andamento)
            print('❌ Não possuí opção de enviar o relatório para o cliente')
            p.press('esc', presses=5)
            return False
            
        envia_experiencia()
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Relatório enviado']), nome=andamento)
        print('✔ Relatório enviado')

    p.press('esc', presses=5)
    time.sleep(2)
    return True
    

def pega_empresas_com_exp():
    with fitz.open(os.path.join('ignore', 'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf')) as pdf:
        # Definir os padrões de regex
        padraozinho_nome1 = re.compile(r'Local\n(\d) - (.+)\n')
        padraozinho_nome2 = re.compile(r'Local\n(\d\d) - (.+)\n')
        padraozinho_nome3 = re.compile(r'Local\n(\d\d\d) - (.+)\n')
        padraozinho_nome4 = re.compile(r'Local\n(\d\d\d\d) - (.+)\n')
        prevtexto_nome = ''
        
        # para cada página do pdf
        if os.path.exists(os.path.join('ignore', 'Dados.csv')):
            os.remove(os.path.join('ignore', 'Dados.csv'))
        for page in pdf:
            andamento = f'Pagina = {str(page.number + 1)}'
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                matchzinho_nome = padraozinho_nome1.search(textinho)
                if not matchzinho_nome:
                    matchzinho_nome = padraozinho_nome2.search(textinho)
                    if not matchzinho_nome:
                        matchzinho_nome = padraozinho_nome3.search(textinho)
                        if not matchzinho_nome:
                            matchzinho_nome = padraozinho_nome4.search(textinho)
                            if not matchzinho_nome:
                                continue

                # Guardar o nome da empresa
                matchtexto_nome = matchzinho_nome.group(2)
                # Guardar o código da empresa no DPCUCA
                matchtexto_cod = matchzinho_nome.group(1)

                if matchtexto_nome == prevtexto_nome:
                    continue
                    
                escreve_dados(matchtexto_cod, matchtexto_nome)
                prevtexto_nome = matchtexto_nome
            except:
                _escreve_relatorio_csv(andamento, nome='Erros')
                continue

    empresas = open_lista_dados(os.path.join('ignore', 'Dados.csv'))
    
    return empresas


def envia_experiencia():
    while not _find_img('publicar_doc.png', conf=0.9):
        _click_img('enviar_arquivo.png', conf=0.9)
        time.sleep(1)
    
    if not _find_img('pasta_pessoal_outros.png', conf=0.9):
        _click_img('drop.png', conf=0.9)
        time.sleep(0.5)
        p.press('down')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)
    
    p.hotkey('Alt', 'c')
    time.sleep(2)


@_time_execution
def run():
    andamentos = 'Experiência a Vencer'
    novo = p.confirm(title='Script incrível', text='Gerar nova planilha de dados?', buttons=('Sim', 'Não'))
    
    if novo == 'Não':
        empresas = _open_lista_dados()
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is None:
            return False
    else:
        index = 0
        gera_arquivo(andamentos)
        empresas = pega_empresas_com_exp()
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
    
        if not _login(empresa, andamentos):
            continue
        gera_arquivo(andamentos, cod=empresa[0], cnpj=empresa[1], nome=empresa[2])


if __name__ == '__main__':
    run()
