# -*- coding: utf-8 -*-
import time, fitz, re, os, shutil
from bs4 import BeautifulSoup
from requests import Session

from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _open_lista_dados, _where_to_start, _escreve_relatorio_csv, _escreve_header_csv, _indice


def salvar_arquivo(nome, response):
    try:
        arquivo = open(os.path.join('execução/Notas', nome), 'wb')
    except FileNotFoundError:
        os.makedirs('execução/Notas', exist_ok=True)
        arquivo = open(os.path.join('execução/Notas', nome), 'wb')
    
    for parte in response.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()
    print('✔ Nota Fiscal salva')


def download(empresas, index):
    # Abre o site do SIGISS pronto para validar as notas
    with Session() as s:
        s.get('https://valinhos.sigissweb.com/validarnfe')

    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        num, chave = empresa

        '''try:'''
        query = {'codigo': chave,
                 'operacao': 'validar'}

        # Faz a validação das notas
        response = s.post('https://valinhos.sigissweb.com/nfecentral?oper=validanfe&codigo={}&tipo=V'.format(chave), data=query)
        nome_nota = 'nfe_' + chave + '.pdf'
        # Salva o PDF da nota
        salvar_arquivo(nome_nota, response)
        # analisa a nota e coleta informações
        time.sleep(0.5)
        analisa_nota(nome_nota)
            
        """except:
            # Se der erro para validar a nota
            print('❌ ERRO, nota: ' + chave)
            _escreve_relatorio_csv(';'.join(["'" + chave, "Erro ao baixar a nota, chave xml inválida"]), 'Erros')
            continue"""


def analisa_nota(nome_nota):
    # Abrir o pdf
    arq = os.path.join('execução', 'Notas', nome_nota)
    
    with fitz.open(arq) as pdf:
        
        # Para cada página do pdf, se for a segunda página o script ignora
        for count, page in enumerate(pdf):
            if count == 1:
                continue
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)

                cnpj_cliente = re.compile(r'(.+)\nTELEFONE / FAX').search(textinho).group(1)
                nome_cliente = re.compile(r'(.+)\nDOCUMENTO VALIDADO COM SUCESSO').search(textinho).group(1)
                uf_cliente = re.compile(r'UF\n(.+)').search(textinho).group(1)
                municipio_cliente = re.compile(r'MUNICÍPIO\n.+\n.+\n(.+)').search(textinho).group(1)
                endereco_cliente = re.compile(r'ENDEREÇO(\n.+){5}\n(.+)').search(textinho).group(1)
                numero_nota = re.compile(r'(.+)\nNÚMERO').search(textinho).group(1)
                serie_nota = re.compile(r'ELETRÔNICA DE\nSERVIÇO\n.+\n(.+)').search(textinho).group(1)
                data_nota = re.compile(r'(.+)\n.+\nDATA EMISSÃO').search(textinho).group(1)
                situacao = 'regular ou cancelada'
                acumulador = '6102'
                if municipio_cliente == 'Valinhos':
                    cfps = '9101'
                else:
                    cfps = '9102'
                valor_servico = re.compile(r'VALOR BRUTO DA NOTA FISCAL\n.+\n.+\nR\$ (.+)').search(textinho).group(1)
                valor_descontos = '0,00'
                valor_contabil = valor_servico
                base_de_calculo = valor_servico
                aliquota_iss = re.compile(r'(\d,\d\d)(.+\n){7}ALIQUOTA ISS\(%\)').search(textinho).group(1)
                if re.compile(r'O ISS NÃO DEVE SER RETIDO'):
                    valor_iss = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                    iss_retido = '0,00'
                else:
                    valor_iss = '0,00'
                    iss_retido = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                    
                try:
                    valor_irrf = re.compile(r'R\$ (.+)\n.+\nIRRF').search(textinho).group(1)
                except:
                    valor_irrf = '0,00'
                try:
                    valor_pis = re.compile(r'R\$ (.+)\n.+\nPIS').search(textinho).group(1)
                except:
                    valor_pis = '0,00'
                try:
                    valor_cofins = re.compile(r'R\$ (.+)\n.+\nCOFINS').search(textinho).group(1)
                except:
                    valor_cofins = '0,00'
                try:
                    valor_csll = re.compile(r'R\$ (.+)\n.+\nCSLL').search(textinho).group(1)
                except:
                    valor_csll = '0,00'
                
                valor_crf = '0,00'
                valor_inss = '0,00'
                codigo_do_iten = '17.18'
                quantidade = '1,00'
                valor_unitario = valor_servico
                
                _escreve_relatorio_csv(';'.join([cnpj_cliente,
                                                 nome_cliente,
                                                 uf_cliente,
                                                 municipio_cliente,
                                                 endereco_cliente,
                                                 numero_nota,
                                                 serie_nota,
                                                 data_nota,
                                                 situacao,
                                                 acumulador,
                                                 cfps,
                                                 valor_servico,
                                                 valor_descontos,
                                                 valor_contabil,
                                                 base_de_calculo,
                                                 aliquota_iss,
                                                 valor_iss,
                                                 iss_retido,
                                                 valor_irrf,
                                                 valor_pis,
                                                 valor_cofins,
                                                 valor_csll,
                                                 valor_crf,
                                                 valor_inss,
                                                 codigo_do_iten,
                                                 quantidade,
                                                 valor_unitario]),
                                       nome='Dados das notas')
            
            except():
                print(textinho)
            
    shutil.move(arq, os.path.join('execução', 'Notas', f'nfe_{numero_nota}.pdf'))

@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    download(empresas, index)


if __name__ == '__main__':
    run()
