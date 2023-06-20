# -*- coding: utf-8 -*-
import os, time
from sys import path

path.append(r'modulos')
import guias_sitac, guias_sinticom, guias_sindeepres, guias_sindpd, guias_sinthojur, guias_seaac, guias_secriopreto, guias_sinthoresca

path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice


@_time_execution
def run():
    os.makedirs('execução/Boletos', exist_ok=True)
    # seleciona a lista de dados
    empresas = _open_lista_dados()
    
    # configura de qual linha da lista começar a rotina
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # para cada linha da lista executa
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod_sindicato, cnpj, valor_boleto, valor_recolhido, valor_remuneracao, data_referencia, usuario, senha, funcionarios, responsavel, email = empresa
        # printa o indice da lista
        _indice(count, total_empresas, empresa)
        
        # dicionário de funções, onde cada arquivo referente a execução de um sindicato está vinculado a um número que é o código do sindicato
        sindicatos = {
            '3': '',
            '8': guias_sitac.run,
            '10': '',
            '11': guias_sinticom.run,            #ok
            '16': guias_sindeepres,
            '17': '',
            '19': '',
            '21': '',
            '22': '',
            '23': '',
            '25': '',
            '28': guias_sindpd.run,
            '39': guias_sinthojur.run,           #ok
            '49': '',
            '58': '',
            '65': '',
            '69': '',
            '70': '',
            '100': guias_seaac.run,
            '131': guias_secriopreto,
            '133': '',
            '135': '',
            '148': guias_sinthoresca,
            '162': '',
            '223': ''
        }
        
        # armazena o resultado retornado da função chamada através do dicionário
        resultado = sindicatos[cod_sindicato](empresa)
        resultado = resultado.replace(' - ', ';')
        
        _escreve_relatorio_csv(f'{cod_sindicato};{cnpj};{valor_boleto};{resultado[2:]}', nome='Boletos Sindicato')
        print(resultado)


if __name__ == '__main__':
    run()