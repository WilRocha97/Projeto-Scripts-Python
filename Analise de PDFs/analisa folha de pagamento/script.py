# -*- coding: utf-8 -*-
from sys import path
import fitz, re, os

path.append(r'..\..\_comum')
from comum_comum import time_execution, ask_for_dir, escreve_relatorio_csv, escreve_header_csv


def analisa_folha(file):
	textos, values = [], ['0,00'] * 2
	regex_cnpj = re.compile(r'\d{5}\n([\d\/\-\.]{14,18})')
	regex_inss = re.compile(r'Contr. Social\n\s+([\d\.,]+)')
	regex_fgts = re.compile(r'Base IRRF(\n\s+([\d\.,]+)){4}\n')
	
	pdf = fitz.open(file)
	for page in pdf:
		texto = page.getText('text', flags=1+2+8)

		match = regex_cnpj.search(texto)
		if match:
			cnpj = match.group(1)
			continue

		if 'Eventos de Funcionários' in texto:
			match =  regex_fgts.search(texto)
			if match:
				values[0] = match.group(1).strip()
			else:
				print('>>> Não achou fgts na pagina', page.number)

		match = regex_inss.search(texto)
		if match: values[1] = match.group(1)
	pdf.close()

	texto = f'{cnpj};{values[0]};{values[1]}'
	values = [i.replace('.', '').replace(',', '.') for i in values]
	total = ';{:.2f}'.format(float(values[0]) + float(values[1]))

	return texto + total.replace('.', ',')


@time_execution
def run():
	docs_path = ask_for_dir()
	if not docs_path: return False

	for file in os.listdir(docs_path):
		if file[-4:].lower() != '.pdf': continue
		abs_path = os.path.join(docs_path, file)

		texto = analisa_folha(abs_path)
		if texto: escreve_relatorio_csv(texto)

	escreve_header_csv('cnpj;fgts mes;base inss;total')


if __name__ == '__main__':
 	run()