# -*- coding: utf-8 -*-
from datetime import datetime
from pandas import read_csv
from time import sleep
from sys import path

import pyautogui as a
import os, re, shutil, sys

sys.path.append(r'..\..\_comum')
from pyautogui_comum import find_img, click_img, wait_img

# variaveis globais
_title = 'obtem extrato caixa'


def mover_arquivo(arq, local, pasta):
    old = os.path.join(local, arq)
    new = os.path.join(pasta, arq)
    shutil.move(old, new)


def salva_registro(arq, local):
    cnpj = arq.split('EXT_FINS_RESC')[0]
    arquivo = open(os.path.join(local, 'erro.csv'), 'a')
    arquivo.write(f'{cnpj};Extrato não emitido\n')
    arquivo.close()


def download_arquivos(pasta):
    hoje = datetime.now().strftime('%d/%m/%Y')
    data = a.prompt('Digite a data da mensagem', title=_title , default=hoje)

    aux = ('cnpj', 'certificado')
    empresas = read_csv(r'ignore\solicitadas.csv', header=None, names=aux, sep=';', encoding='latin-1')
    empresas = empresas.groupby(['certificado'])['cnpj'].apply(list)

    for cert, lista_cnpj in empresas.items():
        a.alert(f'Abra o site com o certificado {cert} (Internet Explore).', title='Informativo')
        aux = a.confirm('Iniciar a rotina?', title=_title, buttons=('Sim', 'Não'))
        if aux != 'Sim':
            print('>>> Rotina cancelada')
            break

        click_img('btn_caixa.png', conf=0.9, clicks=2)
        if not wait_img('menu_msg.png', conf=0.9): continue

        for i in ('btn_entrada', 'menu_entrada'):
            click_img(f'{i}.png', conf=0.9)
            sleep(0.5)

        # VERIFICAR SE ESTA SELECIONANDO A OPÇÃO "EXT_RESCIS"
        a.press(('down', 'down', 'down', 'down', 'tab'), interval=0.25)
        a.write(data.replace('/', ''))

        a.press(('tab', 'tab', 'space'), interval=0.3)
        sleep(1.2)

        if not find_img('tbl_mensgem.png', conf=0.9):
            aux = a.confirm('Continuar a rotina?', title=_title, buttons=('Sim', 'Não'))
            if aux != 'Sim':
                print('>>> Rotina cancelada')
                continue

        if find_img('msg_info.png', conf=0.9):
            a.click('imgs/btn_nao.png')
            sleep(0.3)

        # Seleciona todos os checkbox da tela
        a.click(193, 374)
        sleep(0.3)

        if not find_img('btn_check.png', conf=0.9):
            aux = a.confirm('As mensagens estão marcadas?', title=_title, buttons=('Sim', 'Não'))
            if aux != 'Sim': continue

        a.moveTo(630, 461)
        while not find_img('btn_receber.png', conf=0.9):
            a.scroll(-450)
            sleep(0.2)

        click_img('btn_receber.png', conf=0.9)
        wait_img('msg_salvar.png', conf=0.9, timeout=-1)

        a.press('left', interval=0.3)
        a.write(f'{pasta}\\', interval=0.1)
        while find_img('msg_salvar.png', conf=0.9):
            a.press('enter')
            sleep(0.4)

            if find_img('msg_alerta.png', conf=0.9):
                a.press(['left', 'enter'], interval=0.5)
            sleep(0.6)

            if not click_img('msg_salvar.png', conf=0.9):
                aux = a.confirm('Existem mais arquivos para salvar?', title=_title, buttons=('Sim', 'Não'))
                if aux != 'Sim': 
                    break

        print(f'>>> Finalizado certificado {cert}\n')


def verifica_ocorrencia(arq, ext, local):
    with open(os.path.join(local, arq), 'r', encoding='latin-1') as f:
        linhas = f.readlines()

    if ext == '.txt':
        regex = re.compile(r'\d{2}\/\d{4}')
        for i, linha in enumerate(linhas):
            linha = linha.strip()
            if linha.startswith("COMPETENCIAS NAO LOCALIZADAS"):
                if regex.search(linhas[i+1]):
                    return True
            elif linha.startswith("NAO FORAM EMITIDOS OS EXTRATOS"):
                salva_registro(arq, local)
                break
    else:
        for linha in linhas:
            if "DATA_COMPET_1" in linha:
                return True
            elif "MENS_SEM_EXT" in linha:
                salva_registro(arq, local)
                break

    return False


def separa_arquivos(local):
    pasta = os.path.join(local, 'sem ocorrencias')
    os.makedirs(pasta, exist_ok=True)
    for arq in os.listdir(local):
        ext = arq[-4:]
        if ext not in ('.txt', '.rml'): continue

        try:
            if not verifica_ocorrencia(arq, ext, local):
                mover_arquivo(arq, local, pasta)
                ext2 = '.txt' if ext == '.rml' else '.rml'
                arq2 = arq.replace(ext, ext2)
                if os.path.isfile(arq2):
                    mover_arquivo(arq2, local, pasta)
        except FileNotFoundError:
            pass
        except Exception as e:
            print(e)


def run():
    pasta = os.path.join('C:\\', 'recibos')
    os.makedirs(pasta, exist_ok=True)
    download_arquivos(pasta)
    separa_arquivos(pasta)
    print('>>> Arquivos salvos em ', pasta)


if __name__ == '__main__':
    run()
