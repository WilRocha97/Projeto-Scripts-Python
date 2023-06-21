# -*- coding: utf-8 -*-
import requests
from pyautogui import prompt
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _time_execution, _escreve_header_csv, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def consulta(ref, cod_ae, nome, matricula):
    url = 'https://valinhos.strategos.com.br:9096/api/agenciavirtual/faturasPagas?ftMatricula=' + str(matricula[:-1])
    pagina = requests.get(url)
    for fatura in pagina.json():
        referencia = fatura['vwFtReferenciaMesAnoFormatado']
        vencimento = fatura['vwFtVencimentoFormatado']
        emissao = fatura['vwFtEmissaoFormatado']
        valor = fatura['vwFtValorTotal']
        data_pag = fatura['vwFtDataPagamentoFormatado']
        local_pag = fatura['vwFtLocalPagamento']
        
        if str(referencia) == str(ref):
            _escreve_relatorio_csv(f'{cod_ae};{nome};{referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}', nome=nome)
            print(f'✔ {referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}')
            return True
    
    return False
    

@_time_execution
def run():
    ref = prompt(text='Informe a competência')
    nome = 'Faturas DAEV pagas'
    # seleciona a lista de dados
    empresas = _open_lista_dados()
    
    # configura de qual linha da lista começar a rotina
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # para cada linha da lista executa
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod_ae, nome, hidrometro, matricula, senha = empresa
        # printa o indice da lista
        _indice(count, total_empresas, empresa)
        if not consulta(ref, cod_ae, nome, matricula):
            _escreve_relatorio_csv(f'{cod_ae};{nome};Não encontrou fatura paga referente a {ref}' , nome=nome)
            print(f'❌ Não encontrou fatura paga referente a {ref}')
    
    _escreve_header_csv('CÓDIGO AE;NOME;REFERENCIA;VENCIMENTO;EMISSÃO;VALOR;DATA DE PAGAMENTO;LOCAL DE PAGAMENTO', nome=nome)
    
    
if __name__ == '__main__':
    run()
