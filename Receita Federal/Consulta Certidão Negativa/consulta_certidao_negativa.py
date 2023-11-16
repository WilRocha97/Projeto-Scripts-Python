# -*- coding: utf-8 -*-
import tkinter.filedialog
import pyautogui as p
import time
import os
import pyperclip
import pathlib

e_dir = pathlib.Path('execução')


def abrir_chrome(url, fechar_janela=True, outra_janela=False):
    def abrir_nova_janela():
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        if outra_janela:
            while find_img(outra_janela, conf=0.9):
                time.sleep(1)
        
        while not find_img('google.png', conf=0.9):
            time.sleep(1)
            if find_img('restaurar_pagina.png', conf=0.9):
                click_img('restaurar_pagina.png', conf=0.9)
                time.sleep(1)
                p.press('esc')
                time.sleep(1)
    
    if fechar_janela:
        p.hotkey('win', 'm')
    
    if fechar_janela:
        if find_img('chrome_aberto.png', conf=0.99):
            click_img('chrome_aberto.png', conf=0.99, timeout=1)
        else:
            abrir_nova_janela()
    else:
        abrir_nova_janela()
        
        click_img('google.png', conf=0.9)
        time.sleep(1)
        p.hotkey('alt', 'space', 'x')
        time.sleep(1)
    
    for i in range(4):
        p.click(1000, 51)
        time.sleep(1)
        p.hotkey('ctrl', 'a')
        time.sleep(1)
        pyperclip.copy(url)
        time.sleep(1)
        p.hotkey('ctrl', 'v')
        time.sleep(1)
        p.press('delete')
    
    time.sleep(1)
    p.press('enter')


def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    if local == e_dir:
        local = pathlib.Path(local)
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(str(local / f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome}-auxiliar.csv"), 'a', encoding=encode)
    
    f.write(texto + end)
    f.close()


def ask_for_file(title='Abrir arquivo', filetypes='*', initialdir=os.getcwd()):
    root = tkinter.filedialog.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = tkinter.filedialog.askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    return file if file else False


def open_lista_dados(i_dir='ignore', encode='latin-1'):
    ftypes = [('Plain text files', '*.txt *.csv')]
    
    file = ask_for_file(title='Selecione o arquivos com os dados para a consulta', filetypes=ftypes, initialdir=i_dir)
    if not file:
        return False
    
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False
    
    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def where_to_start(idents, encode='latin-1'):
    title = 'Execucao anterior'
    text = 'Deseja continuar execucao anterior?'
    
    res = p.confirm(title=title, text=text, buttons=('sim', 'não'))
    if not res:
        return None
    if res == 'não':
        return 0
    
    ftypes = [('Plain text files', '*.txt *.csv')]
    file = ask_for_file(title='Selecione o arquivo de resumo da execução', filetypes=ftypes)
    if not file:
        return None
    
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return None
    
    try:
        elem = dados[-1].split(';')[0]
        return idents.index(elem) + 1
    except ValueError:
        return 0


def indice(count, total_empresas, empresa):
    if count > 1:
        print(f'[ {len(total_empresas) - (count - 1)} Restantes ]\n\n')
    # Cria um indice para saber qual linha dos dados está
    indice_dados = f'[ {str(count)} de {str(len(total_empresas))} ]'
    
    empresa = str(empresa).replace("('", '[ ').replace("')", ' ]').replace("',)", " ]").replace(',)', ' ]').replace("', '", ' - ')
    
    print(f'{indice_dados} - {empresa}')


def find_img(img, pasta='imgs', conf=1.0):
    path = os.path.join(pasta, img)
    return p.locateOnScreen(path, confidence=conf)


def click_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, button='left', clicks=1):
    img = os.path.join(pasta, img)
    for i in range(timeout):
        box = p.locateCenterOnScreen(img, confidence=conf)
        if box:
            p.click(*box, button=button, clicks=clicks)
            return True
        time.sleep(delay)
    else:
        return False


def verificacoes(consulta_tipo, andamento, identificacao, nome):
    if find_img('inscricao_cancelada.png', conf=0.9):
        escreve_relatorio_csv('{};{};inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB'.format(identificacao, nome), nome=andamento)
        print('❌ inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB')
        return False

    if find_img('NaoFoiPossivel.png', conf=0.9):
        escreve_relatorio_csv('{};{};Não foi possível concluir a consulta'.format(identificacao, nome), nome=andamento)
        print('❌ Não foi possível concluir a consulta')
        return 'erro'

    if find_img('InfoInsuficiente.png', conf=0.9):
        escreve_relatorio_csv('{};{};As informações disponíveis na Secretaria da Receita Federal do Brasil - RFB sobre o contribuinte '
                               'são insuficientes para a emissão de certidão por meio da Internet.'.format(identificacao, nome), nome=andamento)
        print('❌ Informações insuficientes para a emissão de certidão online')
        return False

    if find_img('Matriz.png', conf=0.9):
        escreve_relatorio_csv('{};{};A certidão deve ser emitida para o {} da matriz'.format(identificacao, nome, consulta_tipo), nome=andamento)
        print(f'❌ A certidão deve ser emitida para o {consulta_tipo} da matriz')
        return False
    
    if find_img('cpf_invalido.png', conf=0.9):
        escreve_relatorio_csv('{};{};CPF inválido'.format(identificacao, nome), nome=andamento)
        print(f'❌ CPF inválido')
        return False
    
    if find_img('cnpj_suspenso.png', conf=0.9):
        escreve_relatorio_csv('{};{};CNPJ suspenso'.format(identificacao, nome), nome=andamento)
        print(f'❌ CNPJ suspenso')
        return False
    
    if find_img('declaracao_inapta.png', conf=0.9):
        escreve_relatorio_csv('{};{};CNPJ com situação cadastral declarada inapta pela Secretaria Especial da Receita Federal do Brasil - RFB'.format(identificacao, nome), nome=andamento)
        print(f'❌ Situação cadastral inapta')
        return False
    
    if find_img('ErroNaConsulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)
        p.press('enter')

    return True


def salvar(consulta_tipo, andamento, identificacao, nome):
    # espera abrir a tela de salvar o arquivo
    while not find_img('SalvarComo.png', conf=0.9):
        time.sleep(1)
        if find_img('em_processamento.png', conf=0.9) or find_img('erro_captcha.png', conf=0.9):
            consulta(consulta_tipo, identificacao)

        if find_img('NovaCertidao.png', conf=0.9):
            click_img('NovaCertidao.png', conf=0.9)

        situacao = verificacoes(consulta_tipo, andamento, identificacao, nome)
        
        if not situacao:
            p.hotkey('ctrl', 'w')
            return False
        
        if situacao == 'erro':
            return situacao

    # escreve o nome do arquivo (.upper() serve para deixar em letra maiúscula)
    time.sleep(1)
    p.write(f'{nome.upper()} - {identificacao} - Certidao')
    time.sleep(0.5)
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    erro = 'sim'
    while erro == 'sim':
        try:
            pyperclip.copy('V:\Setor Robô\Scripts Python\Receita Federal\Consulta Certidão Negativa\execução\Certidões ' + consulta_tipo)
            p.hotkey('ctrl', 'v')
            time.sleep(1)
            erro = 'não'
        except:
            print('Erro no clipboard...')
            erro = 'sim'
    p.press('enter')
    time.sleep(1)
    
    p.hotkey('alt', 'l')
    time.sleep(2)
    
    if find_img('substituir.png', conf=0.9):
        p.hotkey('alt', 's')
        
    return True


def consulta(consulta_tipo, identificacao):
    if consulta_tipo == 'CNPJ':
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir'
    else:
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PF/Emitir'
    
    # espera o site abrir e recarrega caso demore
    while not find_img(f'Informe{consulta_tipo}.png', conf=0.9):
        abrir_chrome(link)
        
        for i in range(60):
            time.sleep(1)
            if find_img(f'Informe{consulta_tipo}.png', conf=0.9):
                break
        
    time.sleep(1)

    click_img(f'Informe{consulta_tipo}.png', conf=0.9)
    time.sleep(1)

    p.write(identificacao)
    time.sleep(1)
    p.press('enter')
    return True


def run():
    consulta_tipo = p.confirm(text='Qual o tipo da consulta?', buttons=('CNPJ', 'CPF'))
    pasta = f'Certidões {consulta_tipo}'
    os.makedirs(r'{}\{}'.format(e_dir, pasta), exist_ok=True)
    andamento = f'Certidão Negativa {consulta_tipo}'

    empresas = open_lista_dados()
    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    
    for count, empresa in enumerate(empresas[index:], start=1):
        indice(count, total_empresas, empresa)
        identificacao, nome = empresa
        nome = nome.replace('/', '')
        
        situacao = 'erro'
        # try:
        while situacao == 'erro':
            consulta(consulta_tipo, identificacao)
            situacao = salvar(consulta_tipo, andamento, identificacao, nome)
            p.hotkey('ctrl', 'w')
            
            if not situacao:
                continue
            
            elif situacao:
                print('✔ Certidão gerada')
                escreve_relatorio_csv('{};{};{} gerada'.format(identificacao, nome, andamento), nome=andamento)
                time.sleep(1)
                
        """except:
            time.sleep(2)
            p.hotkey('ctrl', 'w')"""

    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    # try:
    run()
    p.alert(text='Consulta de Certidão Negativa PF e PJ finalizada')
    """except:
        p.alert(text='Consulta de Certidão Negativa PF e PJ finalizada inesperadamente!')"""