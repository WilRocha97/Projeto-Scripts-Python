# -*- coding: utf-8 -*-
import tkinter.filedialog
import os
import time
import pyperclip
import pyautogui as p
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


def wait_img(img, pasta='imgs', conf=1.0, delay=1, timeout=20, debug=False):
    if debug:
        print('\tEsperando', img)

    aux = 0
    while True:
        box = find_img(img, pasta, conf=conf)
        if box:
            return box
        time.sleep(delay)

        if timeout < 0:
            continue
        if timeout == aux:
            break
        aux += 1

    return None


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
    
    
def salvar_pdf(cnpj, nome, debito=''):
    # navega na tela até aparecer o botão de emitir relatório
    while not find_img('emitir_relatorios.png', conf=0.9):
        p.press('pgDn')
    click_img('emitir_relatorios.png', conf=0.9)
    wait_img('salvar.png', conf=0.9, timeout=-1)
    time.sleep(1)
    # escreve o nome do PDF
    pyperclip.copy(f'{nome} - {cnpj} - CNDNI{debito}.pdf')
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)

    # caso já exista um PDF com o mesmo nome ele substitui
    if find_img('salvar_como.png', conf=0.9):
        p.hotkey('alt', 's')
        time.sleep(1)

    texto = f'{cnpj};Com pendências {debito}'
    print(f'❗ Com pendências {debito}')
    escreve_relatorio_csv(texto)

    # esperar aparecer o botão de voltar e clica nele
    wait_img('voltar.png', conf=0.9)
    click_img('voltar.png', conf=0.9)


def consulta_ipva(cnpj, nome):
    # url para entrar no site
    # url_cnpj = 'https://www10.fazenda.sp.gov.br/CertidaoNegativaDeb/Pages/Restrita/PesquisarContribuinte.aspx'

    # espera a pagina inicial para inserir o cnpj
    while not find_img('cnpj.png', conf=0.9):
        if find_img('certificado.png', conf=0.9):
            click_img('certificado.png', conf=0.9)
        if find_img('verificar_impedimentos.png', conf=0.9):
            click_img('verificar_impedimentos.png', conf=0.9)
        time.sleep(1)
        p.click(700, 400)

    click_img('campo.png', conf=0.9)
    time.sleep(1)
    # limpa o campo do cnpj
    p.press('delete', presses=15)
    p.write(cnpj)
    time.sleep(1)
    p.click(700, 400)
    time.sleep(1)
    click_img('consultar.png', conf=0.9)
    time.sleep(2)

    # aguarda a tela de carregamento
    while find_img('aguarde.png', conf=0.9):
        time.sleep(1)

    # espera a tela da empresa abrir e caso apareça a tela de erro da F5 na página
    timer = 0
    while not find_img('dados.png', conf=0.9):
        if find_img('atencao.png', conf=0.9):
            if find_img('atencao_ok.png', conf=0.9):
                click_img('atencao_ok.png', conf=0.9)
            else:
                p.press('enter')
                p.press('f5')
            if find_img('falha.png', conf=0.9):
                p.press('enter')
                p.press('f5')
        
        time.sleep(5)
        timer += 5
        if timer >= 60:
            return False

    time.sleep(1)
    if find_img('atencao_ok.png', conf=0.9):
        click_img('atencao_ok.png', conf=0.9)

    if find_img('falha.png', conf=0.9):
        p.press('enter')
        p.press('f5')
     
    debitos = ('ha_debitos.png', 'ha_pendencias.png', 'ha_pendencias_2.png', 'icms_declarado.png', 'icms_parcelado.png', 'aiim.png', 'ipva.png')
    # navega na tela até aparecer o botão de emitir relatório e caso tenha algum débito, salva o relatório
    while not find_img('emitir_relatorios.png', conf=0.9):
        p.press('pgDn')
        
        for debito in debitos:
            if find_img(debito, conf=0.9):
                while not find_img('emitir_relatorios.png', conf=0.9):
                    p.press('pgDn')
                    if find_img('gia.png', conf=0.9):
                        salvar_pdf(cnpj, nome, debito=' - GIA')
                        return True
                        
                salvar_pdf(cnpj, nome)
                return True

    # define o texto que ira escrever na planilha
    texto = f'{cnpj};Empresa sem pendências'
    print('✔ Empresa sem pendências')
    escreve_relatorio_csv(texto)

    # voltar pra tela de login da empresa
    wait_img('voltar.png', conf=0.9)
    click_img('voltar.png', conf=0.9)
    return True


def run():
    # abre arquivo de dados
    empresas = open_lista_dados()

    # define a primeira empresa que vai executar o script
    index = where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # define o número total de empresas
    total_empresas = empresas[index:]

    # começa a repetição
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, nome = empresa
        nome = nome.replace('/', ' ')
        # printa o indice dos dados
        indice(count, total_empresas, empresa)

        # executa a consulta
        if not consulta_ipva(cnpj, nome):
            return False
    
    return True


if __name__ == '__main__':
    try:
        if not run():
            p.alert(text='Consulta de Certidão Negativa de Débitos Tributários Não Inscritos finalizada inesperadamente!')
        else:
            p.alert(text='Consulta de Certidão Negativa de Débitos Tributários Não Inscritos finalizada')
    except:
        p.alert(text='Consulta de Certidão Negativa de Débitos Tributários Não Inscritos finalizada inesperadamente!')
