# -*- coding: utf-8 -*-
from datetime import datetime
from subprocess import Popen
from pyperclip import paste
from Dados import empresas
from time import sleep
from PIL import Image
import pytesseract as ocr
import pyautogui as a
import os, re, importlib

'''
 Automatiza o processo de escrever, assinar, transmitir a destda sem movimento
 e salvar o comprovante de envio.

 A planilha de dados desse script é um arquivo .py com nome Dados na seguinte estrutura:
 empresas = [
	("cnpj", (("usuario", "senha"), ("contador", "senha")), "inscrição estadual", "razão"),
	...
 ]

 *Antes de executar o script é necessário alterar a variavel '_comp' com a
 competencia.
 
 *Eventualmente o SEDIF vai congelar, quando isso acontecer um erro ocorrerá no modulo 
 de procura de imagem do pyautogui. Para isso ou por qualquer outra situacao que faca
 o script parar entre a etapa de escrita do documento até a parte de transmissão do documento, 
 deve-se fechar o sedif, abrir novamente, apagar o documento criado, fechar novamente e
 executar a partir da empresa que ocorreu o problema.

 *("08694073000213", "BELA LOLOCA CONFECCAO LTDA  0002") empresa de MG não passa pelo SEDIF

 *Planilha gerada por script.
'''

_comp = '022022'
path_tesseract = 'C:\\','Program Files','Tesseract-OCR','tesseract.exe'
ocr.pytesseract.tesseract_cmd = os.path.join(*path_tesseract)

def escrever_relatorio_csv(texto, end='\n'):
	try:
		f = open('Resumo.csv', 'a')
	except:
		f = open('Resumo_erro.csv', 'a')
	f.write(texto + end)
	f.close()


def wait_img(img, conf=1):
	print('Esperando', img)
	while True:
		try:
			box_img = a.locateOnScreen(os.path.join('imgs', img), confidence=conf)
			if box_img: break
			for i in range(3): sleep(1)
		except:
			r = a.confirm("Excecao", buttons=["Continua", "Parar"])
			if r == "Parar": raise Exception
			sleep(0.5)
	return box_img


def click_img(img, conf=1, tempo=0, no_log=False):
	textos = (f'Clicar {img}\n', 'Não localizou\n') if not no_log else ('', '')
	print(textos[0], end='')
	try:
		x, y = a.locateCenterOnScreen(os.path.join('imgs', img), confidence=conf)
		a.click(x, y)
		sleep(tempo)
	except:
		print(textos[1], end='')
		return False
	return True


def click_center():
	x, y = a.size()
	a.click(x/2, y/2)
	sleep(0.2)


def focus(titulo):
	for title in a.getAllTitles():
		if not title.startswith(titulo): continue
		win = a.getWindowsWithTitle(title)
		break
	else:
		print('Nenhuma janela encontrada')
		return False

	if len(win) > 1:
		print('Mais de uma janela correspondente ao titulo')
		return False

	win[0].activate()
	win[0].maximize()
	sleep(0.2)
	return True


def read_values():
	print('Lendo valores')
	texts = []
	img_path = os.path.join('imgs', 'temp_img.png')
	imgs_config = (
		('new_doc_cnpj.png', 35, 145),
		('new_doc_ie.png', 100, 100)
	)

	for config in imgs_config:
		img, left_px, width = config
		r_img = wait_img(img, conf=0.9)
		if not r_img: return False
		
		r_img = (r_img.left + left_px, r_img.top, width, r_img.height)
		a.screenshot(img_path, region=r_img)
		text = ocr.image_to_string(Image.open(img_path))
		texts.append(''.join(i for i in text if i.isdigit()))

	return texts if texts else False


def write_new_doc(razao, cnpj, ie):
	print('Escrever documento')
	for dado in (razao, _comp):
		a.press('tab')
		sleep(0.2)
		a.write(dado)
		sleep(0.2)

	valores = read_values() 
	if not valores:
		print('Sem valores')
		return False
	if valores[0] != cnpj and valores[1] != ie:
		print(f">>>>>>>{valores} [{cnpj}, {ie}]<<<<<<<")
		escrever_relatorio_csv(f'{razao};{cnpj};Empresa errada')
		return False
	
	a.hotkey('alt', 'p')
	wait_img('unlock_form.png', conf=0.9)
	a.press(['down', 'tab', 'tab', 'down', 'tab', 'down', 'down'], interval=0.2)
	# preenche parte de escrituração do form - 'destda', 'original', 'sem dados'

	a.hotkey('alt', 'c')
	while True:
		for i in range(3): sleep(1)
		try:
			if a.locateOnScreen(os.path.join('imgs', 'open_icon.png'), confidence=0.9):
				return True
			if a.locateOnScreen(os.path.join('imgs', 'new_doc_erro.png'), confidence=0.9):
				escrever_relatorio_csv(f'{razao};{cnpj};Base já existe')
				a.press('enter')
				return False
		except OSError as ex:
			with open(f"{cnpj} erro.txt", 'w') as f:
				f.write(ex.__str__())
			importlib.reload(a)
		except Exception:
			with open(f"{cnpj} erro.txt", 'w') as f:
				f.write(ex)


def send_doc(acesso):
	sleep(0.5)
	a.hotkey('alt', 'i')
	wait_img('send_form.png', conf=0.9)

	for dado in acesso:
		sleep(0.5)
		a.write(dado)
		a.press('tab')

	while True:
		a.moveTo(10, 10) # necessario, tira mouse de cima do botao
		if click_img('ok_send_form.png', conf=0.9): break
		sleep(0.5)

	wait_img('status_send.png', conf=0.9)
	if a.locateOnScreen(os.path.join('imgs','send_sucess.png'), confidence=0.8):
		return True
	else:
		return False


def close_sedif(process):
	a.hotkey('alt', 'f4')
	sleep(0.2)
	a.hotkey('alt', 's')
	process.wait()
	print('Encerrado')


def emite_destda(cnpj, acessos, ie, razao):
	print('>>>Abrir')
	process = Popen(['C:\SimplesNacional\SEDIF\SEDIF.exe'])
	wait_img('att_screen.png', conf=0.9)

	a.hotkey('alt', 'n')
	if not focus('SEDIF-SN'): return False
	
	print('>>>Escrita')
	a.hotkey('alt', 'n')
	while not a.locateOnScreen(os.path.join('imgs', 'new_doc_screen.png'), confidence=0.9):
		for i in range(3): sleep(1)
		click_img('new_doc_bug_screen.png', no_log=True)
	sleep(0.2)

	if not write_new_doc(razao, cnpj, ie):
		a.hotkey('alt', 'f')
		wait_img('cancel_new_doc.png', conf=0.9)
		a.press(['right', 'enter'], interval=0.2)
		close_sedif(process)
		return False

	print('>>>Geração - a partir de agora se der algum erro excluir o documento criado no sedif')
	a.hotkey('alt', 'b')
	wait_img('open_doc.png', conf=0.9)

	a.hotkey('alt', 'r')
	a.hotkey('alt', 'g')
	wait_img('generate_icon.png', conf=0.9)

	a.hotkey('alt', 'i')
	for i in range(2):
		wait_img('confirm_cert.png', conf=0.9)
		a.hotkey('alt', 'n')
		wait_img('generate_sucess.png', conf=0.9)
		a.press('enter')
	sleep(0.5)

	print('>>>Transmissão')
	a.hotkey('alt', 't')
	wait_img('send_icon.png', conf=0.9)

	for acesso in acessos:
		if acesso[1] == "": continue
		if send_doc(acesso):
			a.press('enter')
			while not click_img('show_doc.png', conf=0.9, no_log=True):
				for i in range(3): sleep(1)

			wait_img('save_doc.png', conf=0.9)
			a.write(f'destda {razao} - {cnpj} - {_comp}.pdf')
			a.hotkey('alt', 'l')
			a.press('enter', presses=2, interval=0.5)
			texto = f'{razao};{cnpj};Destda Transmitida'
			break
	else:
		click_center()
		a.hotkey('ctrl', 'c')
		match = re.search(r'\s+\d{12}\s+.*!\s+(.*)', paste())
		texto = f'{razao};{cnpj};{match.group(1)}'

	escrever_relatorio_csv(texto)	
	click_center()

	print('>>>Encerrar -a partir de agora se der algum erro excluir o documento criado na pasta documentos')
	for tecla in ['f', 'i', 'f']:
		a.hotkey('alt', tecla)
		sleep(0.3)

	wait_img('open_icon.png', conf=0.9)
	a.hotkey('alt', 'f')
	close_sedif(process)

	return True


def run():
	inicio = datetime.now()
	lista_empresas = empresas
	total = len(lista_empresas)

	# Lembrete para alterar a competencia
	texto = "Mudou a competencia?\n\n Competencia atual:" + _comp
	if a.confirm(texto, buttons=["ok", "cancel"]) == "cancel":
		return False
	
	for i, empresa in enumerate(lista_empresas, 1):
		print(f'<{empresa[0]} em processo>')
		emite_destda(*empresa)
		print(f'<{i}/{total} concluido>')

	print(datetime.now() - inicio)


if __name__ == '__main__':
	run()
