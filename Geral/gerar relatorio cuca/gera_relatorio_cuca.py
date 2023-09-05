# -*- coding: utf-8 -*-
from datetime import datetime
import fitz, re, os


def escrever_relatorio_csv(nome, texto, end='\n'):
	with open(nome + '.csv', 'a') as f:
		f.write(texto + end)


def get_file():
	path = os.path.join(r'\\VPSRV02', 'dca', 'setor robô', 'relatorios', 'pdf_cuca')
	try:
		files = [os.path.join(path, basename) for basename in os.listdir(path)]
		selected_file = max(files, key=os.path.getctime)
	except FileNotFoundError:
		return None
	print(selected_file)
	return selected_file


def gera_rel_cuca():
	inicio = datetime.now()	
	regex = re.compile(r'(\d{4})\n(.*)\n?(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})')

	with fitz.open(get_file()) as pdf:
		for page in pdf:
			print('pagina', page.number, end=' ')
			texto = page.getText('text')
			match = regex.findall(texto)

			if not match:
				print('nada aqui ', page.number)
				continue

			print('achou', len(match))
			for info in match:
				escrever_relatorio_csv('Dados', ';'.join(info))

	print(datetime.now() - inicio)


if __name__ == '__main__':
	gera_rel_cuca()
