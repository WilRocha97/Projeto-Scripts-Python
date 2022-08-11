# -*- coding: utf-8 -*-
from pyperclip import paste
from time import sleep
import pyautogui as a
import re

'''
 Gera um arquivo .csv com todas as empresas que estão cadastradas no sedif, esse script é necessário,
 pois a ferramenta de busca, para a geracao da destda, utiliza a razao da empresa para a pesquisa.
 Montar a planilha de dados para a execução principal utilizando o index garante que os nomes sempre
 vão ser iguais aos nomes cadastrados no sedif

 Para iniciar é necessário Abrir o Programa SEDIF e navegar para:
 Cadastro de Contribuites -> Contribuintes

 *Não utiliza planilha de dados
 *Apos gerar o index, verificar se existe alguma razao igual, se tiver é necessário alterar
 o cadastro da empresa no SEDIF e o nome dela no index
'''

_regex = r"(\d{14})\s+(\d+)\s+(\d{11})?\s[A-Z]{2}\s+[A-Z].*[a-z\u00C0-\u017F]\s+(.*)"

def escrever_relatorio_csv(texto, end='\n'):
	try:
		f = open('index.csv', 'a')
	except:
		f = open('index_erro.csv', 'a')
	f.write(texto + end)
	f.close()


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


def click_center():
	x, y = a.size()
	a.click(x/2, y/2)


def anotar(valor):
	match = re.search(_regex, valor)
	if not match:
		return "Não anotou"

	print("Anotou", end=" ")
	texto = f"{match.group(1)};{match.group(2)};{match.group(4)}"
	escrever_relatorio_csv(texto)
	return texto


def run():
	focus('SEDIF-SN')
	click_center()

	a.press('end')
	a.hotkey('ctrl', 'c')
	stop = paste().split('Nome Empresa')[1].strip()
	print('stop point :', stop)

	a.press('home')
	sleep(1)

	cont = 0
	while True:
		a.hotkey('ctrl', 'c')
		a.press('down')
		cont += 1
		valor = paste().split('Nome Empresa')[1].strip()
		if valor == stop: break
		print(anotar(valor))

	print(anotar(valor))
	print(cont+1, 'linhas')


if __name__	== '__main__':
	run()