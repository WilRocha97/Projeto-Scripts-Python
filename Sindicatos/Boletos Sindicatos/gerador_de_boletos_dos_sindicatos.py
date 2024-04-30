# -*- coding: utf-8 -*-
import datetime, os, time
from sys import path

path.append(r'modulos')
import guias_sinecol, guias_sitac, guias_metalcampinas, guias_sinticom, guias_sindeepres, guias_sindesporte, guias_sinpospetro, guias_sinsaude, \
    guias_sincomerciariosj, guias_sindquimicos, guias_sindvest_jundiai, guias_sindpd, guias_sindiesp, guias_sinthojur, guias_seectthjr, guias_sinthoresp, \
    guias_secsp, guias_sinditerceirizados, guias_seaac, guias_secriopreto, guias_sindcomerciarios, guias_sincomerciariosr, \
    guias_sinthoresca, guias_sindcargas, guias_sintracargas
    
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
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        cod_sindicato, cnpj, valor_boleto, valor_recolhido, valor_remuneracao, data_referencia, usuario, senha, funcionarios, responsavel, email = empresa
        
        # dicionário de funções, onde cada arquivo referente a execução de um sindicato está vinculado a um número que é o código do sindicato
        sindicatos = {
            '3': guias_sinecol.run,
            '8': guias_sitac.run,
            '10': guias_metalcampinas.run,       #ok
            '11': guias_sinticom.run,            #ok
            '16': guias_sindeepres.run,
            '17': guias_sindesporte.run,
            '19': guias_sinpospetro.run,
            '21': guias_sinsaude.run,
            '22': guias_sincomerciariosj.run,
            '23': guias_sindquimicos.run,
            '25': guias_sindvest_jundiai.run,
            '28': guias_sindpd.run,
            '38': guias_sindiesp.run,
            '39': guias_sinthojur.run,           #ok
            '49': guias_seectthjr.run,
            '58': guias_sinthoresp.run,
            '65': guias_secsp.run,
            '69': guias_sinditerceirizados.run,
            '100': guias_seaac.run,
            '131': guias_secriopreto.run,
            '133': guias_sindcomerciarios.run,
            '135': guias_sincomerciariosr.run,
            '148': guias_sinthoresca.run,
            '162': guias_sindcargas.run,
            '223': guias_sintracargas.run  #223 ou 70
        }
        
        # armazena o resultado retornado da função chamada através do dicionário
        # a função chamada recebe os dados da empresa como parametro
        resultado = sindicatos[str(cod_sindicato)](empresa)
        print(resultado)
        
        # formata o resultado para ser inserido na planilha
        resultado = resultado.replace(' - ', ';')
        _escreve_relatorio_csv(f'{cod_sindicato};{cnpj};{valor_boleto};{resultado[2:]}', nome='Boletos Sindicato')
        


if __name__ == '__main__':
    run()
