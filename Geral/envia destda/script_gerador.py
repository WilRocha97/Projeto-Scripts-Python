# -*- coding: utf-8 -*-
from pandas import read_excel, read_csv
from utilitarios.raw_dados import empresas
from datetime import datetime
import os

'''
 Automatiza o processo de geração da planilha de Dados para o processo
 do script_destda

 A planilha de dados desse script é um arquivo .py com nome Dados na seguinte estrutura:
 empresas = [
	"cnpj",
	...
 ]

 *Antes de executar é necessário alterar a variavel '_comp' com a competencia
 no formato MMAAAA
 *A planilha de dados deve ser salva na pasta 'utilitarios' com o nome 'raw_dados'.py
'''

_comp = "012022"

# cria dataframe ae
def get_info_ae():
	path = os.path.join(r"\\VPSRV02", "dca", "Setor Robô", "relatorios", "ae")

	try:
		files = [os.path.join(path, name) for name in os.listdir(path)]
		selected_file = max(files, key=os.path.getctime)
	except FileNotFoundError:
		selected_file = None

	if not selected_file: return None

	col = {
		"PostoFiscalContador": "contador", "PostoFiscalUsuario": "usuario",
		"PostoFiscalSenha": "senha", "CNPJ": "cnpj",
	}
	data = read_excel(selected_file)
	data.rename(columns=col, inplace=True)
	data.fillna(value={"senha": "", "contador": ""}, inplace=True)
	return data


# cria dataframe index
def get_info_index():
	file = os.path.join("utilitarios", "index.csv")
	if not os.path.exists(file): return None

	cols = ("cnpj", "ie", "razao")
	data = read_csv(file, header=None, names=cols, sep=';', encoding='latin-1')
	return data


def normalize_cont(cont):
	try:
		cont = cont.lower().replace('.', '')
	except:
		print(cont)
		raise Exception

	if any([x in cont for x in ['rodrigo', 'rpem', 'r postal']]):
		return ('RPOSTAL','f7j54kymq4')
	if any([x in cont for x in ['evandro', 'rpmo', 'rp ']]):
		return ('EVMASOL','ty63y26227')
	if any([x in cont for x in ['veiga', 'avj']]):
		return ('VEIGAJR','yq4j8degy5')
	if 'roselei' in cont:
		return ('CNROSELEI','yds34b9p8k')
	return ('', '')


def write_doc(nome, texto):	
	os.makedirs(f"result {_comp}", exist_ok=True)
	path = os.path.join(f"result {_comp}", nome)
	with open(path, "a", encoding='utf-8') as f:
		f.write(texto + "\n")


def run():
	comeco = datetime.now()
	print("Execução iniciada as: ", comeco)
	lista_empresas = empresas

	data_ae = get_info_ae()
	if data_ae is None:
		print("Erro - Não pode encontrar relatorio do AE")
		return True

	data_index = get_info_index()
	if data_index is None:
		print("Erro - Não pode encontrar index.csv")
		return True

	write_doc(f"Dados {_comp}.py", "empresas = [")
	for cnpj in lista_empresas:
		cnpj = "".join(i for i in cnpj if i.isdigit())
		result = data_ae.query("(cnpj == cnpj) and (cnpj == @cnpj)")
		if result.empty:
			texto = f"{cnpj};Não estava no AE"
			write_doc(f"Críticas {_comp}.csv", texto)
			continue

		result = result.iloc[0]
		if not any((result.senha, result.contador)):
			texto = f"{cnpj};Não tem acesso do posto fiscal ou contador"
			write_doc(f"Críticas {_comp}.csv", texto)
			continue

		cont = normalize_cont(result.contador)
		dados = [cnpj, result.usuario, result.senha, *cont]

		result = data_index.query("(cnpj == cnpj) and (cnpj == @cnpj)")
		if result.empty:
			write_doc(f"Sem cadastro {_comp}.csv", cnpj)
			continue

		result = result.iloc[0]
		dados.extend([result.ie, result.razao])
		texto = '\t("{}", (("{}", "{}"), ("{}", "{}")), "{}", "{}"),'.format(*dados)
		write_doc(f"Dados {_comp}.py", texto)

	write_doc(f"Dados {_comp}.py", "]")
	print("Tempo de execução: ", datetime.now() - comeco)


if __name__ == '__main__':
	run()