# -*- coding: utf-8 -*-
import sys

import pyperclip, time, os, pyautogui as p
from time import sleep
from sys import path
from selenium import webdriver
from selenium.webdriver.common.by import By

path.append(r'..\..\_comum')
from comum_comum import _escreve_relatorio_csv
from chrome_comum import _initialize_chrome, _send_input_xpath
from pyautogui_comum import _find_img, _click_img, _click_position_img

imagens = "V:\\_imagens_comum_python\\imgs_comum_dominio"

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Domínio.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.readline()
user = user.split('/')


def _login(empresa, andamentos):
    try:
        cod, cnpj, nome, regime, movimento = empresa
    except:
        try:
            cod, cnpj, nome, regime = empresa
        except:
            try:
                cod, cnpj, nome = empresa
            except:
                cod, cnpj = empresa
                nome = ''
                
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta=imagens, conf=0.9):
        sleep(1)

    p.click(833, 384)
    
    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
        p.press('f8')
        if _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
            break
    
    sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta=imagens, conf=0.9):
        _click_img('codigo.png', pasta=imagens, conf=0.9)
    p.write(cod)
    sleep(3)
    
    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta=imagens, conf=0.9) or _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
        sleep(1)
        if _find_img('sem_parametro.png', pasta=imagens, conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            sleep(1)
            while not _find_img('parametros.png', pasta=imagens, conf=0.9):
                sleep(1)
            p.press('esc')
            sleep(1)
            return False
            
        if _find_img('nao_existe_parametro.png', pasta=imagens, conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não existe parametro cadastrado para esta empresa']), nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            sleep(1)
            p.hotkey('alt', 'n')
            sleep(1)
            p.press('esc')
            sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
                sleep(1)
            while _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
                sleep(1)
            return False
        
        if (_find_img('empresa_nao_usa_sistema.png', pasta=imagens, conf=0.9) or
                _find_img('empresa_nao_usa_sistema_2.png', pasta=imagens, conf=0.9) or
                _find_img('empresa_nao_usa_sistema_3.png', pasta=imagens, conf=0.9)):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não está marcada para usar este sistema']), nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta=imagens, conf=0.9):
                sleep(1)
            while _find_img('trocar_empresa_2.png', pasta=imagens, conf=0.9):
                sleep(1)
            return False
        
        if _find_img('fase_dois_do_cadastro.png', pasta=imagens, conf=0.9) or _find_img('fase_dois_do_cadastro_2.png', pasta=imagens, conf=0.9):
            p.hotkey('alt', 'n')
            sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta=imagens, conf=0.9) or _find_img('conforme_modulo_2.png', pasta=imagens, conf=0.9):
            p.press('enter')
            sleep(1)

        if _find_img('aviso_regime.png', pasta=imagens, conf=0.9) or _find_img('aviso_regime_2.png', pasta=imagens, conf=0.9):
            p.hotkey('alt', 'n')
            sleep(1)

        if _find_img('aviso.png', pasta=imagens, conf=0.9) or _find_img('aviso_2.png', pasta=imagens, conf=0.9):
            p.hotkey('alt', 'o')
            sleep(1)

        if _find_img('erro_troca_empresa.png', pasta=imagens, conf=0.9) or _find_img('erro_troca_empresa_2.png', pasta=imagens, conf=0.9):
            p.press('enter')
            sleep(1)
            p.press('esc', presses=5, interval=1)
            _login(empresa, andamentos)
    
    if not verifica_empresa(cod):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa não encontrada']), nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False
    
    p.press('esc', presses=5)
    sleep(1)

    return True


def verifica_empresa(cod):
    p.click(1000, 500)
    time.sleep(1)
    while True:
        p.click(1258, 82)
        while True:
            try:
                p.hotkey('ctrl', 'c')
                p.hotkey('ctrl', 'c')
                cnpj_codigo = pyperclip.paste()
                break
            except:
                pass
        
        codigo = cnpj_codigo.split('-')
        codigo = str(codigo[-1].strip())
        codigo = codigo.replace(' ', '')
        if codigo != '':
            break
        
    if codigo != cod:
        print(f'Código da empresa: {cod}')
        print(f'Código encontrado no Domínio: {codigo}')
        return False
    else:
        return True
    

def _login_web(usuario=user[0], senha=user[1]):
    if not _find_img('app_controler.png', pasta=imagens, conf=0.99) or not _find_img('app_controler_desfocado.png', pasta=imagens, conf=0.99):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        status, driver = _initialize_chrome(options)
    
        driver.get('https://www.dominioweb.com.br/')
        _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', usuario, driver)
        _send_input_xpath('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', senha, driver)
        driver.find_element(by=By.ID, value='enterButton').click()
        
        abrir_apps = ['abrir_app.png', 'abrir_app_2.png', 'abrir_app_3.png', 'abrir_app_4.png']
        
        print('>>> Aguardando modulos')
        while not _find_img('modulos.png', pasta=imagens, conf=0.9):
            sleep(1)
            for abrir_app in abrir_apps:
                if _find_img(abrir_app, pasta=imagens, conf=0.9):
                    _click_img(abrir_app, pasta=imagens, conf=0.9)

        driver.quit()
        return True
    else:
        if _find_img('app_controler_desfocado.png', pasta=imagens, conf=0.99):
            _click_img('app_controler_desfocado.png', pasta=imagens, conf=0.99, timeout=1)
        else:
            _click_img('app_controler.png', pasta=imagens, conf=0.99, timeout=1)
        sleep(2)
        if _find_img('lista_de_programas.png', pasta=imagens, conf=0.9):
            p.press('right', presses=2, interval=0.5)
            sleep(1)
            p.press('enter')
            

def _abrir_modulo(modulo, usuario=user[2], senha=user[3]):
    if _find_img('inicial.png', pasta=imagens, conf=0.9):
        return True
    
    print(f'>>> Abrindo modulo {modulo.capitalize()}\n')
    while not _find_img('modulos.png', pasta=imagens, conf=0.9):
        sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    sleep(1)
    _click_img('modulo_' + modulo + '.png', pasta=imagens, conf=0.9, button='left', clicks=2)
    
    timer = 0
    while not _find_img('login_modulo.png', pasta=imagens, conf=0.9):
        sleep(1)
        timer += 1
        if timer > 10:
            contador = 1
            with p.hold('alt'):
                if contador == 1:
                    p.press('tab')
                    time.sleep(1)
                if contador == 2:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                if contador == 3:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                if contador == 4:
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(0.1)
                    p.press('tab')
                    time.sleep(1)
                    contador = 0
                contador += 1
        
        if timer > 30:
            while not _find_img('tela_modulos.png', pasta=imagens, conf=0.9):
                contador = 1
                with p.hold('alt'):
                    if contador == 1:
                        p.press('tab')
                        time.sleep(1)
                    if contador == 2:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                    if contador == 3:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                    if contador == 4:
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(0.1)
                        p.press('tab')
                        time.sleep(1)
                        contador = 0
                    contador += 1

            _click_img('tela_modulos.png', pasta=imagens, conf=0.9)
            sleep(1)
            _click_img('modulo_' + modulo + '.png', pasta=imagens, conf=0.9, button='left', clicks=2)
            timer = 0
    
    _click_position_img('insere_usuario.png', '+', pixels_x=120, pasta=imagens, conf=0.9, clicks=2)
    
    sleep(0.5)
    p.press('del', presses=10)
    p.write(usuario)
    sleep(0.5)
    p.press('tab')
    sleep(0.5)
    p.press('del', presses=10)
    p.write(senha)
    sleep(0.5)
    p.hotkey('alt', 'o')
    while not _find_img('onvio.png', pasta=imagens, conf=0.9):
        sleep(1)

    time.sleep(5)
    
    if _find_img('aviso.png', pasta=imagens, conf=0.9):
        p.hotkey('alt', 'o')
    return True


def _salvar_pdf(nome_arquivo=False, abriu_janela=False):
    if not abriu_janela:
        p.hotkey('ctrl', 'd')
    timer = 0
    while not _find_img('salvar_em_pdf.png', pasta=imagens, conf=0.9):
        print('>>> Aguardando tela para salvar o PDF')
        if _find_img('salvar_em_pdf_2.png', pasta=imagens, conf=0.9):
            break
        time.sleep(1)
        timer += 1
        if timer > 30:
            return False
    
    while not _find_img('cliente_c_selecionado.png', pasta=imagens, conf=0.9):
        print('>>> Verificando se o diretório correto para salvar doi selecionado...')
        if _find_img('cliente_m_selecionado.png', pasta=imagens, conf=0.9):
            break
        if _find_img('cliente_m_selecionado_2.png', pasta=imagens, conf=0.9):
            break
        if _find_img('cliente_m_selecionado_3.png', pasta=imagens, conf=0.9):
            break
        if _find_img('cliente_c_selecionado_2.png', pasta=imagens, conf=0.9):
            break
        if _find_img('cliente_c_selecionado_3.png', pasta=imagens, conf=0.9):
            break
        _click_img('botao.png', pasta=imagens, conf=0.9, timeout=1)
        time.sleep(3)

        if _find_img('cliente_c_2.png', pasta=imagens, conf=0.9):
            _click_img('cliente_c_2.png', pasta=imagens, conf=0.9, timeout=1)
        if _find_img('cliente_c.png', pasta=imagens, conf=0.9):
            _click_img('cliente_c.png', pasta=imagens, conf=0.9, timeout=1)
        if _find_img('cliente_m.png', pasta=imagens, conf=0.9):
            _click_img('cliente_m.png', pasta=imagens, conf=0.9, timeout=1)
        if _find_img('cliente_m_2.png', pasta=imagens, conf=0.9):
            _click_img('cliente_m_2.png', pasta=imagens, conf=0.9, timeout=1)
        time.sleep(5)

    """while True:
        p.hotkey('alt', 'n')
        time.sleep(1)
        try:
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            pyperclip.copy(nome_arquivo)
            p.hotkey('ctrl', 'v')
            time.sleep(1)

            if not _find_img('sem_nome_arquivo.png', conf=0.95) or not _find_img('sem_nome_arquivo_2.png', conf=0.95):
                break

        except:
            pass

    time.sleep(1)"""
    p.hotkey('alt', 's')
    
    timer = 0
    while not _find_img('pdf_aberto.png', pasta=imagens, conf=0.9):
        print('>>> Aguardando gerar o PDF...')
        if _find_img('sera_finalizada.png', pasta=imagens, conf=0.9) or _find_img('sera_finalizada_2.png', pasta=imagens, conf=0.9):
            p.press('esc')
            time.sleep(2)
            return False
        
        if _find_img('erro_pdf.png', pasta=imagens, conf=0.9) or _find_img('erro_pdf_2.png', pasta=imagens, conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f4')
            
        if _find_img('substituir.png', pasta=imagens, conf=0.9) or _find_img('substituir_2.png', pasta=imagens, conf=0.9):
            p.hotkey('alt', 'y')
        if _find_img('adobe.png', pasta=imagens, conf=0.9):
            p.press('enter')
        time.sleep(1)
        timer += 1
        if timer > 30:
            p.hotkey('ctrl', 'd')
            while not _find_img('salvar_em_pdf.png', pasta=imagens, conf=0.9):
                if _find_img('salvar_em_pdf_2.png', pasta=imagens, conf=0.9):
                    break
                time.sleep(1)

            if (not _find_img('cliente_c_selecionado.png', pasta=imagens, conf=0.9)
                    or not _find_img('cliente_c_selecionado_3.png', pasta=imagens, conf=0.9)
                    or not _find_img('cliente_m_selecionado.png', pasta=imagens, conf=0.9)
                    or not _find_img('cliente_m_selecionado_2.png', pasta=imagens, conf=0.9)
                    or not _find_img('cliente_m_selecionado_3.png', pasta=imagens, conf=0.9)):
                while not _find_img('cliente_c.png', pasta=imagens, conf=0.9):
                    if _find_img('cliente_c_2.png', pasta=imagens, conf=0.9):
                        break
                    if _find_img('cliente_m.png', pasta=imagens, conf=0.9):
                        break
                    if _find_img('cliente_m_2.png', pasta=imagens, conf=0.9):
                        break
                    _click_img('botao.png', pasta=imagens, conf=0.9)
                    time.sleep(3)
                if _find_img('cliente_c_2.png', pasta=imagens, conf=0.9):
                    _click_img('cliente_c_2.png', pasta=imagens, conf=0.9, timeout=1)
                if _find_img('cliente_c.png', pasta=imagens, conf=0.9):
                    _click_img('cliente_c.png', pasta=imagens, conf=0.9, timeout=1)
                if _find_img('cliente_m.png', pasta=imagens, conf=0.9):
                    _click_img('cliente_m.png', pasta=imagens, conf=0.9, timeout=1)
                if _find_img('cliente_m_2.png', pasta=imagens, conf=0.9):
                    _click_img('cliente_m_2.png', pasta=imagens, conf=0.9, timeout=1)
                time.sleep(5)

            p.press('enter')
            timer = 0

    while _find_img('pdf_aberto.png', pasta=imagens, conf=0.9):
        p.hotkey('alt', 'f4')
        time.sleep(3)
    
    while _find_img('sera_finalizada.png', pasta=imagens, conf=0.9):
        p.press('esc')
        time.sleep(2)

    while _find_img('sera_finalizada_2.png', pasta=imagens, conf=0.9):
        p.press('esc')
        time.sleep(2)
    return True
