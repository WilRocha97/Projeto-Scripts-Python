# -*- coding: utf-8 -*-
from os import listdir, path
from time import sleep
from tkinter.filedialog import askdirectory, Tk
import pyautogui as p
from datetime import datetime, date
import cv2

# 'auto-py-to-exe': para criar o executável e 'Inno setup Compiler': para criar o instalador


# para cada imagem na pasta selecionada, procura na tela se existe aquele usuário selecionado, se tiver,
# pega a coordenada do centro da imagem referente e faz um cálculo para clicar na área do check para desmarcar o usuário
def localiza_autorizados():
    # aguarda alguma subtela que possa estar aberta
    aguarda_sub_telas()
    for imagem in listdir('users'):
        if not imagem == 'Thumbs.db':
            if p.locateOnScreen(path.join('users', imagem)):
                while p.locateOnScreen(path.join('users', imagem)):
                    p.moveTo(p.locateCenterOnScreen(path.join('users', imagem)))
                    local_mouse = p.position()
                    p.click(int(local_mouse[0] - 63), local_mouse[1])


def aguarda_sub_telas():
    while p.locateOnScreen(path.join('imgs', 'tela_parametros.png')) or p.locateOnScreen(path.join('imgs', 'menu_controle.png')) \
            or p.locateOnScreen(path.join('imgs', 'menu_ajuda.png')) or p.locateOnScreen(path.join('imgs', 'usuario_e_senha.png')):
        sleep(1)
    
    tempo = 0
    while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
        sleep(1)
        tempo += 1
        if tempo >= 60:
            return False

    return True


# configura o tempo de espera para atualizar a lista de usuários
def configura_parametro():
    while not p.locateOnScreen(path.join('imgs', 'menu_controle.png')):
        p.hotkey('alt', 'c')
        sleep(1)
    while not p.locateOnScreen(path.join('imgs', 'tela_parametros.png')):
        p.press('p')
        sleep(1)
    p.press('del', presses=5)
    sleep(1)
    p.write('90')
    sleep(1)
    p.hotkey('alt', 'g')
    while p.locateOnScreen(path.join('imgs', 'tela_parametros.png')):
        sleep(2)
        p.hotkey('alt', 'f')


# se for entre segunda ou sexta-feira, antes das 8 da manhã e depois das 17 e 30 da tarde, o robô verifica o horário limite para ser encerrado
def horario():
    hoje = date.today()  # data de hoje
    if not hoje.weekday() in (5, 6):
        if datetime.now() >= datetime.now().replace(hour=17, minute=30, second=0, microsecond=0):
            if datetime.now() <= datetime.now().replace(hour=8, minute=0, second=0, microsecond=0):
                # Verifica o horário
                if datetime.now() >= _hora_limite:
                    return True
                else:
                    return False
    
    return False


def run():
    # espera a tela de carregamento do módulo do domínio fechar
    while p.locateOnScreen(path.join('imgs', 'tela_de_carregamento.png'), confidence=0.9):
        sleep(1)

    # aguarda alguma subtela que possa estar aberta
    if not aguarda_sub_telas():
        return False
    
    # filtra a lista por nome de usuário e clica na coluna do lado para tirar a seleção da coluna de nomes, pois se estiver selecionada o robô se perde
    if p.locateCenterOnScreen(path.join('imgs', 'lista_usuarios.png'), confidence=0.9):
        p.click(p.locateCenterOnScreen(path.join('imgs', 'lista_usuarios.png'), confidence=0.9), button='left', clicks=2)
    if p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9):
        p.click(p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9))

    # aguarda alguma subtela que possa estar aberta
    if not aguarda_sub_telas():
        return False
    
    # seleciona todos os usuários
    if p.locateCenterOnScreen(path.join('imgs', 'lista_usuarios.png'), confidence=0.9):
        p.hotkey('alt', 't')
        sleep(2)
    
    # enquanto não chegar no final da lista procura os usuários que não iram desconectar e tira a seleção dele
    while not p.locateOnScreen(path.join('imgs', 'seta_baixo_limite.png'), confidence=0.9):
        if not p.locateOnScreen(path.join('imgs', 'seta_baixo.png'), confidence=0.9):
            localiza_autorizados()
            break

        # aguarda alguma subtela que possa estar aberta
        if not aguarda_sub_telas():
            return False
        
        # procura o usuário antes de descer a lista
        localiza_autorizados()

        # aguarda alguma subtela que possa estar aberta
        if not aguarda_sub_telas():
            return False
        
        # clica para descer a lista de usuários
        p.moveTo(p.locateCenterOnScreen(path.join('imgs', 'seta_baixo.png'), confidence=0.9))
        p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_baixo_selecionada.png'), confidence=0.9), button='left', clicks=10)
        
        # procura o usuário após descer a lista
        localiza_autorizados()

    # aguarda alguma subtela que possa estar aberta
    if not aguarda_sub_telas():
        return False
    
    # volta para o topo da lista
    if p.locateCenterOnScreen(path.join('imgs', 'seta_cima.png'), confidence=0.9):
        p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_cima.png'), confidence=0.9), button='left', clicks=50)

    # aguarda alguma subtela que possa estar aberta
    if not aguarda_sub_telas():
        return False
    
    # se o botão de desconectar estiver habilitado, clica nele
    if p.locateCenterOnScreen(path.join('imgs', 'desconectar.png'), confidence=0.9):
        p.hotkey('alt', 'd')
        sleep(2)
    
    # caso apareça a tela para confirmar desconexão mesmo com usuários não ociosos
    if p.locateOnScreen(path.join('imgs', 'continuar_desconexao.png')):
        p.hotkey('alt', 'y')
        
    # se a tela de usuário e senha aparecer, insere as informações para desconectar o usuário
    if p.locateOnScreen(path.join('imgs', 'usuario_e_senha.png')):
        p.click(p.locateCenterOnScreen(path.join('imgs', 'usuario.png'), confidence=0.9), button='left')
        p.write(usuario)
        p.click(p.locateCenterOnScreen(path.join('imgs', 'senha.png'), confidence=0.9), button='left')
        p.write(senha)
        sleep(1)
        p.hotkey('alt', 'o')
        sleep(2)
        # caso apareça a tela para confirmar desconexão mesmo com usuários não ociosos
        if p.locateOnScreen(path.join('imgs', 'continuar_desconexao.png')):
            p.hotkey('alt', 'y')
            
    # se der erro de usuário e senha, encerra o script
    sleep(1)
    if p.locateOnScreen(path.join('imgs', 'usuario_ou_senha_invalidos.png'), confidence=0.5):
        return False
    
    return True

    
if __name__ == '__main__':
    _hora_limite = datetime.now().replace(hour=7, minute=30, second=0, microsecond=0)
    # tenta pegar o usuário e a senha de um arquivo txt, se não conseguir marca um erro na variável e encerra o robô
    try:
        f = open('Dados.txt', 'r', encoding='utf-8')
        f = f.read().split('\n')
        usuario = f[0][9:].replace(' ', '')
        senha = f[1][7:].replace(' ', '')
        
        dados_erros = False
        if not usuario:
            dados_erros = True
        if not senha:
            dados_erros = True
    except:
        dados_erros = True
    
    # se conseguir pegar o usuário e a senha no txt segue o processo
    if not dados_erros:
        tempo = 0
        horario_limite = False
        # verifica se existem arquivos na pasta de usuários que não iram desconectar, se sim, continua
        documentos = listdir('users')
        
        if documentos:
            # enquanto a tela de conexões não estiver aberta, espera 1 minuto, se não aparecer, encerra o script
            while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                sleep(1)
                
                if horario():
                    horario_limite = True
                    break
                
                tempo += 1
                if tempo >= 60:
                    tempo = 'inativo'
                    break
            
            if p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                # configura o tempo de espera para atualizar a lista de usuários
                sleep(1)
                p.click(p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9))
                configura_parametro()
            
            # enquanto a tela estiver aberta repete o ciclo, se após 30 segundos não encontrar a tela, o robô é encerrado
            while p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                if horario():
                    horario_limite = True
                    break
                if not run():
                    break
                if horario():
                    horario_limite = True
                    break
                sleep(30)
    
        # alerta de robô finalizado
        if not documentos:
            p.alert(text=f'Diretório dos usuários vazio, robô finalizado.')
        elif tempo == 'inativo':
            p.alert(text=f'Tela de conexões com a banco de dados não encontrada, robô finalizado.')
        elif horario_limite:
            p.alert(text=f'Horário limite atingido, robô finalizado.')
        else:
            p.alert(text=f'Robô finalizado.')
            
    else:
        p.alert(text=f'Erro ao utilizar arquivo de dados, robô finalizado.')
