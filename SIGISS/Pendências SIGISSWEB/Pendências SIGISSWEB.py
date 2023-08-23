# -*- coding: utf-8 -*-
import time
from requests import Session
from bs4 import BeautifulSoup
import sys
import pyautogui as p
import re, os

sys.path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _time_execution, _open_lista_dados, _where_to_start


def login(cnpj, senha):
    with Session() as s:
        # entra no site
        s.get('https://valinhos.sigissweb.com/')
        
        # loga na empresa
        query = {'loginacesso': cnpj,
                 'senha': senha}
        res = s.post('https://valinhos.sigissweb.com/ControleDeAcesso', data=query)
        
        try:
            soup = BeautifulSoup(res.content, 'html.parser')
            soup = soup.prettify()
            # print(soup)
            regex = re.compile(r"'Aviso', '(.+)<br>")
            regex2 = re.compile(r"'Aviso', '(.+)\.\.\.', ")
            try:
                documento = regex.search(soup).group(1)
            except:
                documento = regex2.search(soup).group(1)
            print(f"❌ {documento}")
            return False, documento, s
        except:
            return True, '', s


def consulta(s, cnpj, nome):
        s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=gerarcertidao')

        salvar = s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=imprimirCert&cnpjCpf=' + cnpj)
        # pego o contexto do link referente ao nome original do arquivo
        filename = salvar.headers.get('content-disposition', '')
        
        # aplico o regex para separar o nome do arquivo (Pendencias/Certidao)
        try:
            # crio um regex para obter o nome original do arquivo
            regex = re.compile(r'filename="(.*)\.pdf"')
            documento = regex.search(filename).group(1).replace('ê', 'e').replace('ã', 'a')
        except:
            # se não encontrar o nome do arquivo, procura por alguma mensagem de erro
            # pega o código da página
            soup = BeautifulSoup(salvar.content, 'html.parser')
            soup = soup.prettify()
            
            # procura no código da página a mensagem de erro
            mensagem = re.compile(r"mensagemDlg\((.+)',").search(soup).group(1)
            mensagem = mensagem.replace("','", ", ").replace("'", "")
            print(f"❌ {mensagem}")
            return False, mensagem, s
        
        # se for certidão cria uma pasta para a certidão
        print(f'>>> Salvando {documento}')
        if documento == 'Certidao':
            caminho = os.path.join('execução', 'Certidões', cnpj + ' - ' + nome + ' - Certidão Negativa.pdf')
            os.makedirs('execução\Certidões', exist_ok=True)
            print(f"✔ Certidão")
        
        # se for pendência cria uma pasta para a pendência
        else:
            caminho = os.path.join('execução', 'Pendências', cnpj + ' - ' + nome + ' - Pendências.pdf')
            os.makedirs('execução\Pendências', exist_ok=True)
            print(f"❗ Pendência")
        
        # pega a resposta da requisição que é o PDF codificado, e monta o arquivo.
        arquivo = open(caminho, 'wb')
        for parte in salvar.iter_content(100000):
            arquivo.write(parte)
        arquivo.close()
        
        return True, documento, s


@_time_execution
def run():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        cnpj, senha, nome = empresa
        nome = nome.replace('/', ' ')
        
        # faz login no SIGISSWEB
        execucao, documento, s = login(cnpj, senha)
        if execucao:
            # se fizer login, consulta a situação da empresa
            execucao, documento, s = consulta(s, cnpj, nome)
        
        # escreve os resultados da consulta
        _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}', nome='Pendências SIGISSWEB Valinhos')
        s.close()
        
        
if __name__ == '__main__':
    run()
