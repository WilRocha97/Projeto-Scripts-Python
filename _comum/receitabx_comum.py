# -*- coding: utf-8 -*-
from pyperclip import paste
from time import sleep
import pyautogui as a
import os

from pyautogui_comum import wait_img, focus

# variáveis globais
comum_paths = (
	r'C:\Program Files (x86)\Programas RFB\Receitanet BX',
	r'C:\receitanet',
)


def open_receitanet_bx():
	for path in comum_paths:
		if not os.path.exists(path):
			continue
		
		for file in os.listdir(path):
			conds = (
				file.startswith('receitanetbx-gui-'),
				file.endswith('.jar')
			)
			if not all(conds):
				continue

			os.startfile(os.path.join(path, file))
			wait_img('screen_proc.png')
			return focus('Receitanet BX', ignore_erro=True)
		else:
			continue
	
	return False, 'Receitanet bx não esta instalado ou não esta nos caminhos conhecidos'


def new_session_bx(cnpj, cert):
	print('>>> Nova sessao', cert)

	res = open_receitanet_bx()
	if not res[0]:
		return res

	# selecionar certificado
	a.press('tab')
	sleep(0.5)
	a.hotkey('shift', 'tab')
	sleep(0.2)
	
	prev_cert = ''
	while True:
		a.hotkey('ctrl', 'c')
		text = paste()
		try:
			text = text.split(':')[1].split()[0]
		except:
			text = '00000000000000'

		if cert == text:
			break
		if prev_cert == text:
			return False, 'proc não encontrado'

		prev_cert = text
		a.press('down')

	# fazer primeiro login
	print('\n>>> Logando', cnpj)
	for key in ('tab', 'tab', 'down', 'tab', 'down', 'tab'):
		a.press(key)
		sleep(0.3)

	a.write(cnpj)
	sleep(0.2)
	a.hotkey('alt', 'e')
	sleep(2)
	return True, text


def new_login_bx(cnpj):
	print('\n>>> Logando', cnpj)

	a.hotkey('ctrl', 't')
	wait_img('screen_proc.png', conf=0.9)

	a.press('tab', presses=3, interval=0.2)
	a.press('home')
	a.write(cnpj)

	sleep(0.2)
	a.hotkey('alt', 'e')
	sleep(2)
