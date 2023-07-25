# -*- coding: utf-8 -*-
import time, pyperclip, os, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _click_img, _wait_img, _find_img, _click_position_img, _click_position_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Confere IR.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


def login():
    p.hotkey('win', 'm')

    if _find_img('chrome_aberto.png', conf=0.99):
        _click_img('chrome_aberto.png', conf=0.99, timeout=1)
    else:
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        while not _find_img('google.png', conf=0.9):
            time.sleep(1)
            p.moveTo(1163, 377)
            p.click()

    link = 'https://portal.conferironline.com.br/dashboard'

    _click_img('maxi.png', conf=0.9, timeout=1)
    p.click(1100, 51)
    time.sleep(1)
    p.write(link.lower())
    time.sleep(1)
    p.press('enter')

    return 'ok'


def consulta(cpf):
    # aguarda barra de pesquisa
    while not _find_img('barra_de_pesquisa.png', conf=0.9):
        time.sleep(1)
    
    # clica na barra de pesquisa
    _click_img('barra_de_pesquisa.png', conf=0.9)
    time.sleep(1)
    
    # escreve o cpf
    p.write(cpf)
    
    # aguarda o CPF consultado aparecer
    while not _find_img('empresa.png', conf=0.9):
        if _find_img('nenhum_resultado_encontrado.png', conf=0.9):
            print('❗ CPF não encontrado no sistema')
            return 'CPF não encontrado no sistema'
        time.sleep(1)
    
    # abre o perfil do CPF
    _click_img('empresa.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o menu do ecac aparecer
    while not _find_img('acoes_ecac.png', conf=0.9):
        time.sleep(1)
    
    # clica no menu do ecac
    _click_img('acoes_ecac.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o botão de CND no ecac aparecer, se aparecer uma mensagem sobre login inválido retorna o erro
    while not _find_img('cnd_ecac.png', conf=0.9):
        if _find_img('login_invalido.png', conf=0.9):
            print('❗ Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.')
            return 'Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.'
        time.sleep(1)
    
    # clica no botão de CND do ecac
    _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)

    resultado = gera_relatorio()
    if resultado != 'ok':
        return resultado
    
    # aguarda a tela para salvar o PDF
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_sistema_5.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.press('f5')
            time.sleep(2)

            while not _find_img('cnd_ecac.png', conf=0.9):
                time.sleep(1)

            _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)

            resultado = gera_relatorio()
            if resultado != 'ok':
                return resultado

        time.sleep(1)

    time.sleep(2)
    return 'ok'


def gera_relatorio():
    # aguarda o botão de emissão da certidão aparecer
    timer = 0
    while not _find_img('gerar_relatorio.png', conf=0.9):
        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('erro_sistema_3.png', conf=0.9):
            print('❗ Solicitação rejeitada pelo sistema, tente novamente mais tarde.')
            return 'Solicitação rejeitada pelo sistema, tente novamente mais tarde.'

        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('mensagens_importantes_ecac.png', conf=0.9):
            print(
                '❗ Este CPF possuí mensagens importantes no ECAC, não é possível emitir o relatório até que a mensagem seja aberta.')
            return 'Este CPF possuí mensagens importantes no ECAC, não é possível emitir o relatório até que a mensagem seja aberta.'

        # se demorar 5 segundos para o botão de emissão da certidão aparecer, verifica se a tela de login do ecac stá bugada
        # se a tela de login do ecac estiver bugada, fecha a janela, recarrega a página no conferir e clica no botão de CND do ecac novamente
        if timer > 5:
            # se aparecer a tela de, em processamento, retorna o erro
            """if _find_img('consulta_em_processamento.png', conf=0.9):
                print('❗ Consulta em processamento, retorne mais tarde.')
                return 'Consulta em processamento, retorne mais tarde.'"""

            if _find_img('erro_sistema.png', conf=0.9) \
                    or _find_img('erro_sistema_2.png', conf=0.9) \
                    or _find_img('erro_sistema_4.png', conf=0.9)\
                    or _find_img('erro_sistema_6.png', conf=0.9):
                p.hotkey('ctrl', 'w')
                time.sleep(1)
                p.press('f5')
                time.sleep(2)

                while not _find_img('cnd_ecac.png', conf=0.9):
                    time.sleep(1)

                _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)
                timer = 0
        time.sleep(1)
        timer += 1

    time.sleep(2)
    # clica para emitir a certidão
    _click_img('gerar_relatorio.png', conf=0.9)
    _click_img('botao_gerar_relatorio.png', conf=0.9)

    return 'ok'


def salvar_pdf():
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)

    pasta_final = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\execução\Relatórios'
    os.makedirs(pasta_final, exist_ok=True)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)

    erro = 'sim'
    while erro == 'sim':
        try:
            pyperclip.copy(pasta_final)
            p.hotkey('ctrl', 'v')
            erro = 'não'
        except:
            erro = 'sim'

    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(2)
    p.hotkey('ctrl', 'w')

    print('✔ Relatório emitido com sucesso')
    return 'Relatório emitido com sucesso'


@_time_execution
def run():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o índice para localizar em qual empresa está
        _indice(count, total_empresas, empresa)
        cpf, nome = empresa
        
        login()
        resultado = consulta(cpf)
        if resultado == 'ok':
            resultado = salvar_pdf()

        p.hotkey('ctrl', 'w')
        _escreve_relatorio_csv(f'{cpf};{nome};{resultado}')


if __name__ == '__main__':
    run()
