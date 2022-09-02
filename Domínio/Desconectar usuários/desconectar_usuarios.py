# -*- coding: utf-8 -*-
from os import listdir, path
from time import sleep
from tkinter.filedialog import askdirectory, Tk
import pyautogui as p
import cv2


# pega a pasta onde estão as imagens referentes aos usuários que não iram desconectar
def ask_for_dir(title='Selecione a pasta onde estão as imagens dos usuários que não iram desconectar'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False


# para cada imagem na pasta selecionada, procura na tela se existe aquele usuário selecionado, se tiver,
# pega a coordenada do centro da imagem referente e faz um cálculo para clicar na área do check para desmarcar o usuário
def localiza_autorizados():
    for imagem in listdir(documentos):
        if p.locateOnScreen(path.join(documentos, imagem)):
            while p.locateOnScreen(path.join(documentos, imagem)):
                p.moveTo(p.locateCenterOnScreen(path.join(documentos, imagem)))
                local_mouse = p.position()
                p.click(int(local_mouse[0] - 63), local_mouse[1])
                sleep(0.5)


def run():
    # espera a tela de carregamento do módulo do domínio fechar
    while p.locateOnScreen(path.join('imgs', 'tela_de_carregamento.png'), confidence=0.9):
        sleep(1)
    
    # filtra a lista por nome de usuário e clica na coluna do lado para tirar a seleção da coluna de nomes, pois se estiver selecionada o robô se perde
    p.click(p.locateCenterOnScreen(path.join('imgs', 'lista_usuarios.png'), confidence=0.9), button='left', clicks=2)
    p.click(p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9))
    
    # seleciona todos os usuários
    p.hotkey('alt', 't')
    sleep(2)
    
    # enquanto não chegar no final da lista procura os usuários que não iram desconectar e tira a seleção dele
    while not p.locateOnScreen(path.join('imgs', 'seta_baixo_limite.png')):
        # procura o usuário antes de descer a lista
        localiza_autorizados()
        
        p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_baixo.png'), confidence=0.9), button='left', clicks=10)
        
        # procura o usuário após descer a lista
        localiza_autorizados()
    
    # volta para o topo da lista
    p.click(p.locateCenterOnScreen(path.join('imgs', 'seta_cima.png'), confidence=0.9), button='left', clicks=50)
    
    # se o botão de desconectar estiver habilitado, clica nele
    if p.locateCenterOnScreen(path.join('imgs', 'desconectar.png'), confidence=0.9):
        p.hotkey('alt', 'd')
        sleep(1)
        
    if p.locateOnScreen(path.join('imgs', 'usuario_e_senha.png')):
        p.click(p.locateCenterOnScreen(path.join('imgs', 'usuario.png'), confidence=0.9), button='left')
        p.write(usuario)
        p.click(p.locateCenterOnScreen(path.join('imgs', 'senha.png'), confidence=0.9), button='left')
        p.write(senha)
        sleep(1)

    # p.hotkey('alt', 'o')
    p.hotkey('alt', 'c')
    
    
if __name__ == '__main__':
    # pergunta qual a pasta com as imagens dos usuários que não iram desconectar
    tempo = 0
    usuario = None
    senha = None
    documentos = ask_for_dir()
    if documentos:
        usuario = p.prompt(text='Usuário Gerente')
        if usuario:
            senha = p.prompt(text='Senha Gerente')
            if senha:
                # espera abrir a tela
                p.alert(text=f'Abra a tela de conexões com o banco de dados e configure o tempo de atualização para 30 segundos (Controle > Parâmetros).')
                while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.5):
                    sleep(1)
                    tempo += 1
                    if tempo >= 60:
                        break
                    
                # enquanto a tela estiver aberta repete o ciclo, se após 30 segundos não encontrar a tela, o robô é encerrado
                while p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.5):
                    run()
                    sleep(30)
    
    # alerta de robô finalizado
    if not documentos:
        p.alert(text=f'Diretório dos usuários não selecionado, robô finalizado.')
    elif not usuario:
        p.alert(text=f'Usuario Gerente não informado, robô finalizado.')
    elif not senha:
        p.alert(text=f'Senha Gerente não informada, robô finalizado.')
    elif tempo == 'inativo':
        p.alert(text=f'Tela de conexões com a banco de dados não encontrada, robô finalizado.')
    else:
        p.alert(text=f'Robô finalizado.')
        