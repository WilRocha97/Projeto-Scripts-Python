# -*- coding: utf-8 -*-
from Dados import empresas
from time import sleep
import pyautogui as a

'''
 Automatiza o processo de cadastro das empresa no sedif

 Para iniciar é necessário Abrir o Programa SEDIF e navegar para:
 Cadastro de Contribuites -> Novo Contribuinte -> Cancelar

 *Algumas informações do cadastro foram pré setadas no script para diminuir o Dados.py
 *Apos cadastrar novas empresas, fazer um index novo para manter o index integro
 *Todas as informações para montar a planilha de dados estão no AE

 A planilha de dados desse script é um arquivo .py com nome Dados na seguinte estrutura:
 empresas = [
	("cnpj", "ie", "razao", "uf", "cidade"**, "cep", "rua", "numero", "bairro", "resp"***, "r_cpf"),
   ...
 ]

 *Planilha feita manualmente.
 **campos "cidade" é correspondente ao numero da cidade no sedif
 ***campos "resp" é correspondente ao nome do sócio principal da empresa
'''

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


def run():
	focus('SEDIF-SN')
	click_center()

	#numero de tabs na parte de empresa
	e_tabs = [1, 4, 1, 1, 2, 1, 1, 2, 3]

	#numero de tabs na parte de contador
	c_tabs = [1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 2, 3, 2] 

	#informações pre setadas do contador
	contador = [
		"Aldemar Veiga Junior", "900", "25194605803", "SP189659O7",
		"19937946000107", "veigaepostal@veigaepostal.com.br", "SP",
		"3556206", "13271260", "Rua Fioravante Basilio Maglio", "345",
		"Nova Valinhos", "1938718959"
	]
	
	lista_empresa = empresas
	for empresa in lista_empresa:
		*aux, resp, r_cpf = empresa

		#tratamento de dados - empresa
		aux[0] = "".join(i for i in aux[0] if i.isdigit()).rjust(14, "0")
		aux[1] = "".join(i for i in aux[1] if i.isdigit()).rjust(12, "0")
		aux[2] = "".join(i for i in aux[2] if i not in ["\\/:*?'\"><|"])
		aux[5] = "".join(i for i in aux[5] if i.isdigit()).rjust(8, "0")
		aux[8] = aux[8] if aux[8] else "-"

		#tratamento de dados - resp
		r_cpf = "".join(i for i in r_cpf if i.isdigit()).rjust(11, "0")

		a.hotkey('alt', 'n')
		sleep(1)
		
		#empresa
		for i, dado in zip(e_tabs, aux):
			a.write(dado)
			a.press("tab", presses=i)
		else:
			a.write(contador[-1]) # telefone pre setado
			a.press("tab", presses=3)

		#responsavel
		a.write(resp)
		a.press("tab")
		a.write("801") # codigo de empresário presetado no sedif
		a.press("tab", presses=2)
		a.write(r_cpf)
		a.press("tab")
		a.write(contador[5]) # email presetado
		a.press("tab")

		#contador
		for i, dado in zip(c_tabs, contador):
			a.write(dado)
			a.press("tab", presses=i)

		a.hotkey('alt', 'c') # confirmar cadastro
		sleep(1.5)


if __name__ == '__main__':
	run()