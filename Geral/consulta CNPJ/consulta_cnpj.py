# -*- coding: utf-8 -*-
# from json.decoder import JSONDecodeError
from sys import path
import requests, time, sys

path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _escreve_header_csv, _time_execution, _open_lista_dados, _where_to_start


@_time_execution
def run():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    url = "https://www.receitaws.com.br/v1/cnpj/"

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, numero = empresa
        _indice(count, total_empresas, empresa, index)
        lista_dados = []
        atividades = ''
        socios = ''

        try:
            pagina = requests.get(url + cnpj)
            print('>>> Aguardando resultado da requisição')

            toolbar_width = 50
            # setup toolbar
            sys.stdout.write("[%s]" % (" " * toolbar_width))
            sys.stdout.flush()
            sys.stdout.write("\b" * (toolbar_width + 1))  # return to start of line, after '['
            for i in range(toolbar_width):
                time.sleep(0.422)  # do real work here
                # update the bar
                sys.stdout.write("■")
                sys.stdout.flush()
            sys.stdout.write("]\n")  # this ends the progress bar

        except Exception as e:
            print('\t❌ Erro durante a requisição: ', e)
            continue
        try:
            dados = pagina.json()
        except Exception as e:
            print('\t❌ Erro durante a requisição: ', e)
            continue

        if dados['status'] == 'ERROR':
            print(f"❌ Ocorreu um erro com o CNPJ {cnpj}: {dados['message']}")
            _escreve_relatorio_csv(f"{cnpj};{dados['message']}", 'ocorrencias')
            continue

        lista_dados.append(dados['cnpj'])
        lista_dados.append(dados['nome'])
        lista_dados.append(dados['fantasia'])
        lista_dados.append(dados['situacao'])
        lista_dados.append(dados['abertura'])
        lista_dados.append(dados['natureza_juridica'])
        lista_dados.append(dados['tipo'])
        lista_dados.append(dados['porte'])
        lista_dados.append(dados['atividade_principal'][0]['code'])
        lista_dados.append(dados['logradouro'])
        lista_dados.append(dados['numero'])
        lista_dados.append(dados['complemento'])
        lista_dados.append(dados['bairro'])
        lista_dados.append(dados['uf'])
        lista_dados.append(dados['cep'])
        lista_dados.append(dados['municipio'])
        lista_dados.append(dados['telefone'])
        lista_dados.append(dados['email'])
        texto = ';'.join(lista_dados)

        atividades_secundarias = dados['atividades_secundarias']
        for a in atividades_secundarias:
            atividades += ';' + a['code']
        qsa = dados['qsa']
        for q in qsa:
            socios += ';' + q['qual'] + ';' + q['nome']

        print('✔ Resultado obtido')
        _escreve_relatorio_csv(cnpj + ';' + texto + atividades, 'atividades')
        _escreve_relatorio_csv(texto + socios, 'socios')

    lista_cabecalho = [
        'CNPJ', 'Nome', 'Fantasia', 'Situação', 'Abertura', 'Natureza Jurídica', 'Tipo', 'Porte',
        'Atividade Princ.', 'Logradouro', 'Número', 'Compl.', 'Bairro', 'UF', 'CEP',
        'Município', 'Telefone', 'Email'
    ]

    cabecalho = ";".join(lista_cabecalho)
    _escreve_header_csv(cabecalho + ';Atividade Secund.', 'atividades.csv')
    _escreve_header_csv(cabecalho + ';Qualificação;Nome Sócio', 'socios.csv')


if __name__ == '__main__':
    run()
