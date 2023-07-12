# -*- coding: utf-8 -*-
from os import listdir, path
from time import sleep
from tkinter.filedialog import askdirectory, Tk
import cv2, pyautogui as p
from datetime import datetime, date


# 'auto-py-to-exe': para criar o executável e 'Inno setup Compiler': para criar o instalador

# para cada imagem na pasta selecionada, procura na tela se existe aquele usuário selecionado, se tiver,
# pega a coordenada do centro da imagem referente e faz um cálculo para clicar na área do check para desmarcar o usuário

def captura_dados():
    f = open('Dados.txt', 'r', encoding='utf-8')
    f = f.read().split('\n')
    usuario = f[0][8:].replace(' ', '')
    senha = f[1][6:].replace(' ', '')
    refresh_timer = f[2][17:].replace(' ', '')
    
    hora = f[3][8:].replace(' ', '')
    hora_iniciar = hora.split(':')

    if hora_iniciar[0] == '0':
        inicio_horario = 0
        fim_horario = 0
    else:
        inicio_horario = datetime.now().replace(hour=int(hora_iniciar[0]), minute=int(hora_iniciar[1]), second=0, microsecond=0)
        
        hora = f[4][7:].replace(' ', '')
        hora_finalizar = hora.split(':')
        fim_horario = datetime.now().replace(hour=int(hora_finalizar[0]), minute=int(hora_finalizar[1]), second=0, microsecond=0)
        
    dados_erros = False
    if not usuario:
        dados_erros = True
        usuario = ''
    if not senha:
        dados_erros = True
        senha = ''
    
    return refresh_timer, inicio_horario, fim_horario, dados_erros, usuario, senha


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
                    sleep(0.1)


def aguarda_sub_telas():
    while p.locateOnScreen(path.join('imgs', 'tela_parametros.png')) or p.locateOnScreen(path.join('imgs', 'menu_controle.png')) \
            or p.locateOnScreen(path.join('imgs', 'menu_ajuda.png')) or p.locateOnScreen(path.join('imgs', 'usuario_e_senha.png')):
        sleep(0.1)
    
    tempo = 0
    while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
        sleep(0.1)
        tempo += 0.1
        if tempo >= 600:
            return False
    
    return True


# configura o tempo de espera para atualizar a lista de usuários
def configura_parametro(refresh_timer):
    if refresh_timer == '':
        refresh_timer = '60'
    while not p.locateOnScreen(path.join('imgs', 'menu_controle.png')):
        p.hotkey('alt', 'c')
        sleep(1)
    while not p.locateOnScreen(path.join('imgs', 'tela_parametros.png')):
        p.press('p')
        sleep(1)
    p.press('del', presses=5)
    sleep(1)
    p.write(refresh_timer)
    sleep(1)
    p.hotkey('alt', 'g')
    while p.locateOnScreen(path.join('imgs', 'tela_parametros.png')):
        sleep(2)
        p.hotkey('alt', 'f')


# se for entre segunda ou sexta-feira, verifica o horário limite para ser encerrado
def horario():
    refresh_timer, inicio_horario, fim_horario, dados_erro, usuario, senha = captura_dados()
    if inicio_horario != 0:
        while True:
            refresh_timer, inicio_horario, fim_horario, dados_erro, usuario, senha = captura_dados()
            timer = 0
            while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                sleep(1)
                timer += 1
                if timer > 60:
                    return False
            
            now = datetime.now()
            if inicio_horario <= now <= fim_horario:
                return True
            else:
                print("Aguardando horário personalizado...")
                sleep(60)  # Espera por 1 minuto antes de verificar novamente
    else:
        return True


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
        p.hotkey('alt', 'n')
        sleep(2)
        
    # enquanto não chegar no final da lista procura os usuários que não iram desconectar e tira a seleção dele
    if not p.locateOnScreen(path.join('imgs', 'seta_baixo.png'), confidence=0.9):
        localiza_autorizados()
    else:
        while not p.locateOnScreen(path.join('imgs', 'seta_baixo_limite.png'), confidence=0.9):
            if not horario():
                return False
                
            # aguarda alguma subtela que possa estar aberta
            if not aguarda_sub_telas():
                return False
            
            # procura o usuário antes de descer a lista
            localiza_autorizados()
            
            # clica para descer a lista de usuários
            p.click(p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9), button='left')
            p.press('pgdn')
            
            if p.locateOnScreen(path.join('imgs', 'seta_baixo_limite.png'), confidence=0.9):
                localiza_autorizados()
    
    # aguarda alguma subtela que possa estar aberta
    if not aguarda_sub_telas():
        return False
        
    # volta para o topo da lista
    p.click(p.locateCenterOnScreen(path.join('imgs', 'estacao.png'), confidence=0.9), button='left')
    p.press('pgup', presses=15)
    
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
    # tenta pegar o usuário e a senha de um arquivo txt, se não conseguir marca um erro na variável e encerra o robô
    try:
        refresh_timer, inicio_horario, fim_horario, dados_erro, usuario, senha = captura_dados()
    except:
        refresh_timer = ''
        dados_erro = True
    
    # se conseguir pegar o usuário e a senha no txt segue o processo
    if not dados_erro:
        tempo = 0
        # verifica se existem arquivos na pasta de usuários que não iram desconectar, se sim, continua
        documentos = listdir('users')
        
        if documentos:
            # enquanto a tela de conexões não estiver aberta, espera 1 minuto, se não aparecer, encerra o script
            while not p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                sleep(1)
                
                if not horario():
                    tempo = 'inativo'
                    break
                
                tempo += 1
                if tempo >= 60:
                    tempo = 'inativo'
                    break
            
            if p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                # configura o tempo de espera para atualizar a lista de usuários
                sleep(1)
                p.click(p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9))
                configura_parametro(refresh_timer)
            
            # enquanto a tela estiver aberta repete o ciclo, se após 30 segundos não encontrar a tela, o robô é encerrado
            while p.locateOnScreen(path.join('imgs', 'conexoes.png'), confidence=0.9):
                if not horario():
                    tempo = 'inativo'
                    break
                    
                situacao = run()
                if not situacao:
                    break
                
                if not horario():
                    tempo = 'inativo'
                    break
                    
                timer = 1
                while timer < 30:
                    sleep(1)
                    timer += 1
                    horario()
        
        # alerta de robô finalizado
        if not documentos:
            p.alert(text=f'Diretório dos usuários vazio, robô finalizado.')
        elif tempo == 'inativo':
            p.alert(text=f'Tela de conexões com a banco de dados não encontrada, robô finalizado.')
        else:
            p.alert(text=f'Robô finalizado.')
    
    else:
        p.alert(text=f'Usuário e/ou senha não informados, robô finalizado.')
