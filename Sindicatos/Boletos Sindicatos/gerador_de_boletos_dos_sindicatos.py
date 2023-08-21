# -*- coding: utf-8 -*-
import os, time
from sys import path

path.append(r'modulos')
import guias_sinecol, guias_sitac, guias_metalcampinas, guias_sinticom, guias_sindeepres, guias_sindesporte, guias_sincomerciarios, \
    guias_sindpd, guias_sinthojur, guias_sinthoresp, guias_secsp, guias_seaac, guias_secriopreto, guias_sinthoresca

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
        _indice(count, total_empresas, empresa, index)
        
        # dicionário de funções, onde cada arquivo referente a execução de um sindicato está vinculado a um número que é o código do sindicato
        sindicatos = {
            '3': guias_sinecol.run,
            '8': guias_sitac.run,
            '10': guias_metalcampinas.run,       #ok
            '11': guias_sinticom.run,            #ok
            '16': guias_sindeepres.run,
            '17': guias_sindesporte.run,
            '19': guias_sinpospetro.run,
            '21': '',
            '22': guias_sincomerciarios.run,
            '23': '',
            '25': '',
            '28': guias_sindpd.run,
            '39': guias_sinthojur.run,           #ok
            '49': '',
            '58': guias_sinthoresp.run,
            '65': guias_secsp.run,
            '69': '',
            '70': '',
            '100': guias_seaac.run,
            '131': guias_secriopreto.run,
            '133': '',
            '135': '',
            '148': guias_sinthoresca.run,
            '162': '',
            '223': guias_sintracargas.run
        }
        
        # armazena o resultado retornado da função chamada através do dicionário
        resultado = sindicatos[str(cod_sindicato)](empresa)
        resultado = resultado.replace(' - ', ';')
        
        _escreve_relatorio_csv(f'{cod_sindicato};{cnpj};{valor_boleto};{resultado[2:]}', nome='Boletos Sindicato')
        print(resultado)


if __name__ == '__main__':
    run()