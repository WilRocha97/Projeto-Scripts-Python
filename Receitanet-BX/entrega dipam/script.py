# -*- coding: utf-8 -*-
from pyautogui import prompt
from time import sleep
from sys import path
import pyautogui as a
import os, shutil

path.append(r'..\..\_comum')
from receitabx_comum import new_login_bx, new_session_bx
from pyautogui_comum import find_img, wait_img, click_img, get_comp
from comum_comum import escreve_relatorio_csv, open_lista_dados, where_to_start, time_execution


def solicitar(dt_i, dt_f):
	print('>>> pesquisa e solicitação')
	dt_i, dt_f = dt_i.split('/'), dt_f.split('/')
	dt_i = tuple([dt_i[0][0], dt_i[0][1]] + dt_i[1:])
	dt_f = tuple([dt_f[0][0], dt_f[0][1]] + dt_f[1:])

	
	click_img('btn_search.png', conf=0.9)
			
	sleep(0.2)

	# sistema
	for i in range(5):
		a.hotkey('shift', 'tab')
	a.press('down', presses=5)
	
	# tipo arquivo
	for key in ('tab', 'down'):
		a.press(key)
		sleep(0.2)

	# dt inicio e dt final
	a.press('down', presses=3)
	for i in dt_i:
		a.write(i)
	a.press('down')
	for i in dt_f:
		a.write(i)
	a.press('down')
	sleep(1)

	# marcar opcao ultimo arquivo
	click_img('check_result_alt.png', conf=0.9)
	a.press('down')

	a.hotkey('ctrl', 'p')  # pesquisar
	while True:
		if find_img('search_results.png', conf=0.9):
			sleep(1)
			break

		if find_img('info_proc.png', conf=0.9):
			a.press('enter')
			return 'Erro de procuração'

		if find_img('info_proc_cancel.png', conf=0.9):
			a.press('enter')
			return 'Erro de procuração'

		if find_img('info_erro.png', conf=0.9):
			a.press('enter')
			return 'Erro de procuração'

		if find_img('none_result.png', conf=0.9):
			a.press('enter')
			return 'Nenhum resultado'

	click_img('check_result_baixar.png', conf=0.9)
	sleep(0.2)

	a.hotkey('ctrl', 's')
	wait_img('screen_request.png', conf=0.9)
	a.press('enter')


def baixar():
	print('>>> download do pedido')
	while True:
		if click_img('btn_queue.png', conf=0.9):
			break
	sleep(1)
	
	click_img('ver_pedidos.png', conf=0.9)
		
	# seleciona arquivo do topo da lista
	a.press('tab', presses=8, interval=0.2)
	wait_img('screen_files.png', conf=0.9)

	# marca o arquivo para download
	a.press('tab', presses=14, interval=0.1)
	click_img('check_result_baixar.png', conf=0.9)
	sleep(0.2)

	# baixar arquivo e esperar concluir
	for i in range(2):
		click_img('btn_download.png', conf=0.9)
		sleep(0.2)

	while True:
		if find_img('empty_queue.png', conf=0.9):
			print('>>> movendo arquivo')

			src = os.path.join(os.path.expandvars("%userprofile%"), 'documents', 'arquivos receitanetbx')
			dest = os.path.join(os.getcwd(), 'execucao', 'docs')
			os.makedirs(dest, exist_ok=True)

			for file in os.listdir(src):
				shutil.move(os.path.join(src, file), dest)

			texto = 'arquivo baixado'
			break

		if find_img('info_exists.png', conf=0.9):
			a.press('enter')
			texto = 'arquivo já foi baixado'
			break

	a.press('tab', presses=12, interval=0.1)
	a.press('right')
	return texto


@time_execution
def run():
	prev_cert = ''

	empresas = open_lista_dados()
	if not empresas:
		return False

	index = where_to_start(tuple(i[0] for i in empresas))
	if index is None:
		return False

	dt_i = get_comp(printable='dd/mm/yyyy', strptime='%d/%m/%Y', subject='dt inicio')
	if not dt_i:
		return False

	dt_f = get_comp(printable='dd/mm/yyyy', strptime='%d/%m/%Y', subject='dt final')
	if not dt_f:
		return False

	for cnpj, cert in empresas[index:]:
		if prev_cert != cert:
			if prev_cert:
				os.system("taskkill /F /IM javaw.exe")
				sleep(0.5)

			status, msg = new_session_bx(cnpj, cert)
			if not status:
				print('>>>', msg)
				return False
			else:
				prev_cert = cert
		else:
			new_login_bx(cnpj)

		# necessário para dar foco a janela certa
		for i in range(2):
			a.hotkey('alt', 'tab')
		sleep(0.5)

		sit = solicitar(dt_i, dt_f)
		if not sit:
			sit = baixar()

		print(f'>>> {sit}')
		escreve_relatorio_csv(f'{cnpj};{sit}')

	os.system("taskkill /F /IM javaw.exe")


if __name__ == '__main__':
	run()
