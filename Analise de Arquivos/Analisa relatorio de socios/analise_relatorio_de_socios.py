# -*- coding: utf-8 -*-
import fitz, re, os, pandas as pd
from datetime import datetime
from time import sleep
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir


def run():
	comeco = datetime.now()
	print(comeco)

	regex_empresa = re.compile(r"(\d{5})([A-Z\u00C0-\u024E\.\d/].+)\n\d{4}\n")
	regex_socio = re.compile(r"\n([A-Z\u00C0-\u024E]+ [A-Z\u00C0-\u024E][A-Z\s\u00C0-\u024E\.]+)\n")
	regex_insc = re.compile(r"Inscrição C\.Indiv\.\n(.*)\n")
	regex_cpf = re.compile(r"([\d\.-]{14})\n CPF")
	
	documentos = ask_for_dir()
	# Analiza cada pdf que estiver na pasta
	for arquivo in os.listdir(documentos):
		print(f'\nArquivo: {arquivo}')
		# Abrir o pdf
		pdf = fitz.open(os.path.join(documentos, arquivo))
	
		dados = []
		pcod, prazao = [""] * 2
	
		for page in pdf:
			texto = page.get_text("text", flags=1+2+8)
	
			match = regex_empresa.search(texto)
			cod, razao = match.groups() if match else (pcod, prazao)
	
			match = regex_socio.search(texto)
			socio = match.group(1) if match else ""
	
			match = regex_cpf.search(texto)
			cpf = match.group(1) if match else ""
	
			insc = ""
			match = regex_insc.search(texto)
			if match:
				insc = [i for i in match.group(1) if i.isdigit()]
				insc = ''.join(insc) if len(insc) > 0 else ""
	
			dados.append((cod, razao, socio, insc, cpf))
			pcod, prazao = cod, razao
		pdf.close()
	
		# Guarda informações em uma planilha do Excel
		data = {
			"codigo" : [i[0] for i in dados],
			"razao": [i[1] for i in dados],
			"nome socio": [i[2] for i in dados],
			"inscricao socio": [i[3] for i in dados],
			"cpf":[i[4] for i in dados],
		}
		columns = ["codigo", "razao", "nome socio", "inscricao socio", "cpf"]
		df = pd.DataFrame(data, columns = columns)
		df.to_excel('resumo.xlsx', index = False)
		# Guarda informações em uma planilha do Excel
	
		print(datetime.now() - comeco, "-" * 60)


if __name__ == '__main__':
	run()