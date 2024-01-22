# -*- coding: utf-8 -*-
import requests
from pyautogui import prompt
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _time_execution, _escreve_header_csv, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def consulta(ref, nome_planilha, cod_ae, cnpj, nome, matricula):
    url = 'https://valinhos.strategos.com.br:9096/api/agenciavirtual/faturasPagas?ftMatricula=' + str(matricula[:-1])
    # fas um get na api usando o número da matrícula sem o último digito
    pagina = requests.get(url)
    # usa a variável 'comp' para saber se é uma consulta anual ou mensal
    comp = ''
    
    for fatura in pagina.json():
        try:
            # com o json capturado, pega as infos desejadas
            referencia = fatura['vwFtReferenciaMesAnoFormatado']
            # separa a data de referência para usar caso a consulta seja anual
            ref_ano = referencia.split('/')[1]
            vencimento = fatura['vwFtVencimentoFormatado']
            emissao = fatura['vwFtEmissaoFormatado']
            valor = fatura['vwFtValorTotal']
            data_pag = fatura['vwFtDataPagamentoFormatado']
            local_pag = fatura['vwFtLocalPagamento']
        except:
            _escreve_relatorio_csv(f'{cod_ae};{nome};{cnpj};{matricula};Água-ID/Matrícula inválida', nome=nome_planilha)
            print(f'❌ Água-ID/Matrícula  inválida')
            return True
    
        # escreve se for consulta mensal e se encontrar alguma conta do mês procurado
        if str(referencia) == str(ref):
            _escreve_relatorio_csv(f'{cod_ae};{nome};{cnpj};{matricula};{referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}', nome=nome_planilha)
            print(f'✔ {referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}')
            return True
        
        # escreve se for consulta anual e todas as contas do ano procurado se existir
        if str(ref_ano) == str(ref):
            comp = 'anual'
            _escreve_relatorio_csv(f'{cod_ae};{nome};{cnpj};{matricula};{referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}', nome=nome_planilha)
            print(f'✔ {referencia};{vencimento};{emissao};{valor};{data_pag};{local_pag}')
    
    if comp == '':
        return False
    if comp == 'anual':
        return True
    

@_time_execution
def run():
    ref = prompt(text='Informe a competência (00/0000 para mês | 0000 para ano)')
    ref_planilha = ref.replace('/', '-')
    nome_planilha = f'Faturas DAEV pagas ref. {str(ref_planilha)}'
    # seleciona a lista de dados
    empresas = _open_lista_dados()
    
    # configura de qual linha da lista começar a rotina
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # para cada linha da lista executa
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cod_ae, cnpj, nome, matricula = empresa
        # printa o indice da lista
        _indice(count, total_empresas, empresa, index)
        if not consulta(ref, nome_planilha, cod_ae, cnpj, nome, matricula):
            _escreve_relatorio_csv(f'{cod_ae};{nome};{cnpj};{matricula};Não encontrou fatura paga referente a {ref}', nome=nome_planilha)
            print(f'❌ Não encontrou fatura paga referente a {ref}')
        
    _escreve_header_csv('CÓDIGO AE;NOME;REFERENCIA;MATRÍCULA;REFERENCIA;VENCIMENTO;EMISSÃO;VALOR;DATA DE PAGAMENTO;LOCAL DE PAGAMENTO', nome=nome_planilha)
    
    
if __name__ == '__main__':
    run()
