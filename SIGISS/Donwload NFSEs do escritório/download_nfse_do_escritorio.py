# -*- coding: utf-8 -*-
from requests import Session
from sys import path
import os

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _indice


def salvar_arquivo(nome, response):
    try:
        arquivo = open(os.path.join('execucao/Documentos', nome), 'wb')
    except FileNotFoundError:
        os.makedirs('execucao/Documentos', exist_ok=True)
        arquivo = open(os.path.join('execucao/Documentos', nome), 'wb')
        
    for parte in response.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()


def download(empresas, index):
    # Abre o site do SIGISS pronto para validar as notas
    with Session() as s:
        s.get('https://valinhos.sigissweb.com/validarnfe')

    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        num, chave = empresa

        try:
            query = {'codigo': chave,
                     'operacao': 'validar'}

            # Faz a validação das notas
            response = s.post('https://valinhos.sigissweb.com/nfecentral?oper=validanfe&codigo={}&tipo=V'.format(chave), data=query)

            # Salva o PDF da nota e anota na planilha
            salvar_arquivo('nfe_' + num + '.pdf', response)
            _escreve_relatorio_csv(';'.join([num, 'Nota Fiscal salva']), 'Download NFSEs do Escritório')
            print('✔ Download concluído, nota: ' + num)
        except:
            # Se der erro para validar a nota
            print('❌ ERRO, nota: ' + num)
            _escreve_relatorio_csv(';'.join([num, 'Erro']), 'Download NFSEs do Escritório')
            continue


@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    download(empresas, index)


if __name__ == '__main__':
    run()
