# -*- coding: utf-8 -*-
import datetime, pyautogui as p
import time
from sys import path

path.append(r'..\..\_comum')
from comum_comum import _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login

_login_web()
_abrir_modulo('escrita_fiscal', usuario='ROBO2', senha='Rb#0086*')


@_barra_de_status
def run(window):
    contador = 0
    
    horarios = [('17:50', True), ('22:00', True), ('23:59', False)]
    horarios_2 = [('02:00', True), ('04:00', True), ('06:00', False)]

    for horario in horarios:
        while True:
            p.click(600, 750)
            p.click(800, 750)
            time.sleep(1)
            # Defina o horário após o qual você deseja executar a condição
            horario_especifico = datetime.datetime.strptime(horario[0], "%H:%M").time()
            # Obtenha a data e hora atuais
            horario_atual = datetime.datetime.now().time()
            # Verifique se o horário atual é igual ou após o horário específico
            if horario_atual >= horario_especifico:
                break

        print(datetime.datetime.now())
        _login_web()
        _abrir_modulo('escrita_fiscal')
        contador += 1
        window['-Mensagens-'].update(f'Reinícios: {contador}')
        window.refresh()

        if not horario[1]:
            for horario_2 in horarios_2:
                while True:
                    p.click(600, 750)
                    p.click(800, 750)
                    time.sleep(1)
                    # Defina o horário após o qual você deseja executar a condição
                    horario_especifico = datetime.datetime.strptime(horario_2[0], "%H:%M").time()
                    # Obtenha a data e hora atuais
                    horario_atual = datetime.datetime.now().time()
                    # Verifique se o horário atual é igual ou após o horário específico
                    if horario_atual >= horario_especifico:
                        break

                print(datetime.datetime.now())
                _login_web()
                _abrir_modulo('escrita_fiscal')
                contador += 1
                window['-Mensagens-'].update(f'Reinícios: {contador}')
                window.refresh()
    
    
if __name__ == '__main__':
    run()
