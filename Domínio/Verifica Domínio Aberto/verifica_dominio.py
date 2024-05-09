# -*- coding: utf-8 -*-
import datetime, psutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login

_login_web()
_abrir_modulo('escrita_fiscal', usuario='ROBO2', senha='Rb#0086*')


@_barra_de_status
def run(window):
    contador = 0
    
    horarios = [('17:50', True), ('23:59', False)]
    
    for horario in horarios:
        # Defina o horário após o qual você deseja executar a condição
        horario_especifico = datetime.datetime.strptime(horario[0], "%H:%M").time()
        
        # Variável para rastrear se a condição já foi executada
        condicao_executada = False
        
        # Loop infinito para verificar o horário atual
        while not condicao_executada:
            p.moveTo(600,750)
            p.click(600,750)
            p.moveTo(800, 750)
            p.click(800, 750)
            # Obtenha a data e hora atuais
            horario_atual = datetime.datetime.now().time()
            
            # Verifique se o horário atual é igual ou após o horário específico
            if horario_atual >= horario_especifico:
                # Execute a condição
                _login_web()
                _abrir_modulo('escrita_fiscal', usuario='ROBO2', senha='Rb#0086*')
                contador += 1
                window['-Mensagens-'].update(f'Reinícios: {contador}')
                window.refresh()
                
                # Defina a variável para True para encerrar o loop
                condicao_executada = horario[1]
        
                if not horario[1]:
                    while True:
                        if _find_img('reconnect.png', pasta='imgs_c', conf=0.9):
                            _click_img('reconnect_nao.png', pasta='imgs_c', conf=0.9)

                        p.moveTo(600, 750)
                        p.click(600, 750)
                        p.moveTo(800, 750)
                        p.click(800, 750)
                        if not "AppController.exe" in (i.name() for i in psutil.process_iter()):
                            _login_web()
                            _abrir_modulo('escrita_fiscal')
                            contador += 1
                            window['-Mensagens-'].update(f'Reinícios: {contador}')
                            window.refresh()
        
    
if __name__ == '__main__':
    run()
