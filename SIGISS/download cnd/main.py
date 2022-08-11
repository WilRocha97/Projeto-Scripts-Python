# -*- coding: utf-8 -*-
# ! python3
from requests import Session
import os
import re
from datetime import datetime

'''
Download de CND ou lista de débitos de cada cpf/cnpj listado no arquivo 'lista.txt' 
Os arquivos são obtidos pelo site 'https://valinhos.sigissweb.com'
'''

def escreve_relatorio_csv(texto):
    try:
        arquivo = open("Resumo.csv", 'a')
    except:
        arquivo = open("Resumo.csv", 'w')
    arquivo.write(texto+'\n')
    arquivo.close()


def mover():
    os.mkdir('pendencia', exist_ok=True)
    arquivo = open('Resumo.csv', 'r')
    lista = arquivo.readlines()
    arquivo.close()

    for linha in lista:
        cpf, situacao = linha.strip().split(';')
        if situacao == 'Pendencias':
            old = os.path.join('certidao', f'{cpf}.pdf')
            new = os.path.join('pendencia', f'{cpf}.pdf')
            shutil.move(old, new)


def consulta_cnd():
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    certidao = "https://valinhos.sigissweb.com/CertidaoNegativaCentral?oper=imprimirCert&cnpjCpf="
    login = "24982859000101"
    senha = "2376500"

    # cria a pasta 'certidoes' caso ela não exista
    os.makedirs('certidao', exist_ok=True)

    # obtem a lista de cpf/cnpj do arquivo texto
    lista = open('lista.txt', 'r')
    cpfcnpj = lista.readlines()
    lista.close()

    # inicia a sessão no site, realiza o login e obtem os arquivos para cada contribuinte
    with Session() as s:
        login_data = {"loginacesso":login,"senha":senha}
        pagina = s.post(url_acesso,login_data)
        if pagina.status_code != 200:
            print('>>> Erro ao acessar página.')
            return False
        # crio um regex para obter o nome original do arquivo
        regex = re.compile(r'filename="(.*)\.pdf"')
        for num in cpfcnpj[:]:
            num = num.strip()
            link = certidao + num
            salvar = s.get(link)
            
            if 'text' not in salvar.headers.get('Content-Type'):
                # pego o contexto do link referente ao nome original do arquivo
                filename = salvar.headers.get('content-disposition', '')
                # aplico o regex para separar o nome do arquivo (Pendencias/Certidao)
                nome = regex.search(filename).group(1).replace('ê', 'e').replace('ã', 'a')
                print('>>> Nome -> ', nome)
                caminho = os.path.join('certidao', num+'.pdf')
                arquivo = open(caminho, 'wb')
                for parte in salvar.iter_content(100000):
                    arquivo.write(parte)
                arquivo.close()
                print(f"Arquivo {num}.pdf salvo")
                escreve_relatorio_csv(num+';'+nome)
            else:
                escreve_relatorio_csv(num+';Empresa não localizada')
    return True


if __name__ == '__main__':
    comeco = datetime.now()
    concluido = consulta_cnd()
    mover()
    if concluido:
        print('\n>>> Programa executado com sucesso.')
    else:
        print('\n>>> Falha durante execução do programa.')
    fim = datetime.now()
    print("\n>>> Tempo total:", (fim-comeco))