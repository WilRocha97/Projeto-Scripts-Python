# -*- coding: utf-8 -*-
import time
from requests import Session
from bs4 import BeautifulSoup
import sys
import pyautogui as p
import re, os

sys.path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _time_execution, _open_lista_dados, _where_to_start, _indice


def consulta(empresas, index):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        
        cnpj, senha, nome = empresa
        nome = nome.replace('/', ' ')
        
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
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}', nome='Pendências SIGISSWEB Valinhos')
                print(f"❌ {documento}")
                continue
            except:
                pass
            
            s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=gerarcertidao')

            # crio um regex para obter o nome original do arquivo
            regex = re.compile(r'filename="(.*)\.pdf"')
            salvar = s.get('https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=imprimirCert&cnpjCpf=' + cnpj)
            # if 'text' not in salvar.headers.get('Content-Type'):
            
            # pego o contexto do link referente ao nome original do arquivo
            filename = salvar.headers.get('content-disposition', '')
            
            # aplico o regex para separar o nome do arquivo (Pendencias/Certidao)
            try:
                documento = regex.search(filename).group(1).replace('ê', 'e').replace('ã', 'a')
            except:
                soup = BeautifulSoup(salvar.content, 'html.parser')
                soup = soup.prettify()
                
                mensagem = re.compile(r"mensagemDlg\((.+)',").search(soup).group(1)
                mensagem = mensagem.replace("','", ", ").replace("'", "")
                _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{mensagem}', nome='Pendências SIGISSWEB Valinhos')
                print(f"❌ {mensagem}")
                continue
                
            print(f'>>> Salvando {documento}')
            if documento == 'Certidao':
                caminho = os.path.join('execução', 'Certidões', cnpj + ' - ' + nome + ' - Certidão Negativa.pdf')
                os.makedirs('execução\Certidões', exist_ok=True)
                print(f"✔ Certidão")
            else:
                caminho = os.path.join('execução', 'Pendências', cnpj + ' - ' + nome + ' - Pendências.pdf')
                os.makedirs('execução\Pendências', exist_ok=True)
                print(f"❗ Pendência")
                
            arquivo = open(caminho, 'wb')
            for parte in salvar.iter_content(100000):
                arquivo.write(parte)
            arquivo.close()
            
            _escreve_relatorio_csv(f'{cnpj};{senha};{nome};{documento}', nome='Pendências SIGISSWEB Valinhos')
            

@_time_execution
def run():
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    consulta(empresas, index)


if __name__ == '__main__':
    run()
