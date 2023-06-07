# -*- coding: utf-8 -*-
import tkinter.filedialog
import pyautogui as p
import time
import os
import pyperclip
import pathlib

e_dir = pathlib.Path('execução')


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


def verificacoes(consulta_tipo, andamento, empresa):
    identificacao, nome = empresa
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
    
    if find_img('ErroNaConsulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)
        p.press('enter')

    return True


def salvar(consulta_tipo, andamento, empresa, pasta):
    identificacao, nome = empresa
    # espera abrir a tela de salvar o arquivo
    while not find_img('SalvarComo.png', conf=0.9):
        time.sleep(1)
        if find_img('em_processamento.png', conf=0.9):
            consulta(consulta_tipo, empresa)

        if find_img('NovaCertidao.png', conf=0.9):
            click_img('NovaCertidao.png', conf=0.9)

        situacao = verificacoes(consulta_tipo, andamento, empresa)
        
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
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Receita Federal\Consulta Certidão Negativa\execução\Certidões CPF')
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    
    p.hotkey('alt', 'l')
    time.sleep(2)
    
    if find_img('substituir.png', conf=0.9):
        p.hotkey('alt', 's')
        
    return True


def consulta(consulta_tipo, empresa):
    identificacao, nome = empresa
    p.hotkey('win', 'm')

    # Abrir o site
    if find_img('Chrome.png', conf=0.99):
        pass
    elif find_img('ChromeAberto.png', conf=0.99):
        click_img('ChromeAberto.png', conf=0.99)
    else:
        time.sleep(0.5)
        click_img('ChromeAtalho.png', conf=0.9, clicks=2)
        while not find_img('Google.png', conf=0.9, ):
            time.sleep(5)
            p.moveTo(1163, 377)
            p.click()

    if consulta_tipo == 'CNPJ':
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir'
    else:
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PF/Emitir'

    click_img('Maxi.png', conf=0.9, timeout=1)
    # espera o site abrir e recarrega caso demore
    while not find_img(f'Informe{consulta_tipo}.png', conf=0.9):
        p.click(1150, 51)
        time.sleep(1)
        p.write(link.lower())
        time.sleep(1)
        p.press('enter')
        for i in range(10):
            time.sleep(1)
            if find_img(f'Informe{consulta_tipo}.png', conf=0.9):
                break
            
    while not find_img(f'Informe{consulta_tipo}.png', conf=0.9):
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
        situacao = 'erro'
        try:
            while situacao == 'erro':
                consulta(consulta_tipo, empresa)
                situacao = salvar(consulta_tipo, andamento, empresa, pasta)
                
                if not situacao:
                    p.hotkey('ctrl', 'w')
                    continue
                
                elif situacao:
                    identificacao, nome = empresa
                    print('✔ Certidão gerada')
                    escreve_relatorio_csv('{};{};{} gerada'.format(identificacao, nome, andamento), nome=andamento)
                    time.sleep(1)
                
        except:
            time.sleep(2)
            p.hotkey('ctrl', 'w')

    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    try:
        run()
        p.alert(text='Consulta de Certidão Negativa PF e PJ finalizada')
    except:
        p.alert(text='Consulta de Certidão Negativa PF e PJ finalizada inesperadamente!')
