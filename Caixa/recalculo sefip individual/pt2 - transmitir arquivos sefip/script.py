# -*- coding: utf-8 -*-
from time import sleep
import pyautogui as a
import os, sys

sys.path.append(r'..\..\..\_comum')
from pyautogui_comum import click_img, wait_img, find_img, focus

# variaveis globais
_base = os.path.join(r'\\vpsrv02','dca','sefip', 'arquivos')
_path_sfp = os.path.join(_base, 'sfp')
_path_zip = os.path.join(_base, 'zip')
_path_pdf = os.path.join(_base, 'pdf')


def salvar_arquivos_finais(tipo):
    img = f'btn_salvar_{tipo}.png'
    if not click_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')
    sleep(1)

    # gambiarra - o mouse precisa mover um pouco
    # para apertar as teclas do for, é um bug do site
    a.moveTo(10,10)

    for key in ('tab', 'tab', 'down', 'down', 'enter'):
        a.press(key)
        sleep(0.2)

    img = 'icon_new_file.png'
    if not wait_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')

    a.press('home')
    sleep(0.6)

    a.write(os.path.join(_path_pdf, '@').replace('@', ''))
    a.press('enter')
    sleep(1)

    a.press(('tab', 'tab', 'tab', 'tab', 'enter'))
    print(f'\tArquivo {tipo} salvo')


def gera_zip(file):
    print('\n>>> gerando', file)

    # clica no botão nova mensagem caso esteja disponivel
    click_img('btn_new_msg.png', conf=0.9, timeout=3)

    # preenche o campo municipio
    img = 'inp_municipio.png'
    if not click_img(img, conf=0.9):
        return f'Não encontrou {img}'

    a.hotkey('ctrl', 'a')
    sleep(0.6)
    a.press('delete')
    sleep(0.6)

    a.write('campinas')
    sleep(0.6)

    a.press('tab')
    sleep(0.6)

    # Anexa o arquivo sfp
    img = 'btn_anexar.png'
    if not click_img(img, conf=0.9):
        # apaga o conteudo do campo municipio caso dê errado
        a.hotkey('alt', 'tab')
        sleep(0.6)
        a.press('delete')
        return f'Não encontrou {img}'

    img = 'icon_new_folder.png'
    if not wait_img(img, conf=0.9):
        return f'Não encontrou {img}'

    a.write(os.path.join(_path_sfp, file))
    a.hotkey('alt', 'a')

    for i in range(5):
        if find_img('icon_erro.png', conf=0.9):
            raise Exception('Erro anexando arquivo sfp')
        sleep(0.5)

    print('\tArquivo sfp anexado')

    img = 'btn_salvar.png'
    if not click_img(img, conf=0.9):
        return f'Não encontrou {img}'

    img = 'icon_new_folder.png'
    if not wait_img(img, conf=0.9):
        return f'Não encontrou {img}'

    a.write(os.path.join(_path_zip, f'{file[:-4]}.zip'))
    a.press('enter')

    img = 'icon_erro.png'
    if not wait_img(img, conf=0.9):
        return f'Não encontrou {img}'

    a.press('enter')
    return 'Arquivo zip salvo'


def envia_zip(file):
    print('\n>>> enviando', file)

    # clica no botão enviar caso esteja disponivel
    click_img('btn_envia.png', conf=0.9, timeout=3)

    img = 'btn_procura.png'
    if not click_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')

    img = 'icon_new_file.png'
    if not wait_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')

    a.write(os.path.join(_path_zip, file))
    a.press('enter')

    img = 'btn_envia_zip.png'
    if not click_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')

    img = 'msg_sucesso.png'
    if not click_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')
    print('\tArquivo zip enviado')

    salvar_arquivos_finais('xml')
    sleep(1)
    salvar_arquivos_finais('pdf')
    sleep(1)
    a.hotkey('alt', 'f4')

    img = 'btn_retornar.png'
    if not click_img(img, conf=0.9):
        raise Exception(f'Não encontrou {img}')

    return 'concluido'


def run():
    res = focus('Conectividade Social - Internet Explorer')
    if not res[0]: raise Exception(res[1])

    os.makedirs(_path_zip, exist_ok=True)
    zips = tuple(x[:-4] for x in os.listdir(_path_zip) if x[-3:] == 'zip')

    for file in os.listdir(_path_sfp):
        if file[:-4] in zips: continue
        click_img('btn_new_msg.png', conf=0.9, timeout=3)
        
        result = gera_zip(file)
        print(f'\t{result}')
    else:
        print('>>>Todos os arquivos zips salvos')

    sleep(5)
    os.makedirs(_path_pdf, exist_ok=True)
    xmls = tuple(x[:-4] for x in os.listdir(_path_pdf) if x[-3:] == 'xml')
    pdfs = tuple(x[:-4] for x in os.listdir(_path_pdf) if x[-3:] == 'pdf')

    for file in os.listdir(_path_sfp):
        if 'SEFIP' + file[:-4] in pdfs: continue
        if 'SEFIP' + file[:-4] in xmls:
            os.remove(os.path.join(_path_pdf, f'SEFIP{file[:-4]}.xml'))

        file = file.replace('.SFP', '.zip')
        result = envia_zip(file)
        print(f'\t{result}')
    else:
        print('>>>Todos os arquivos pdfs salvos')


if __name__ == '__main__':
    run()