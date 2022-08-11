# -*- coding: utf-8 -*-
from email.mime.text import MIMEText
from bs4 import BeautifulSoup
from requests import Session
from Dados import empresas
import smtplib
import sys
sys.path.append("..")
from comum import login, time_execution


def salva_registro(cnpj, texto):
    try:
        arquivo = open('ocorrencias.csv', 'a')
    except:
        arquivo = open('ocorrencias.csv', 'w')
    registro = cnpj+';'+texto+'\n'
    arquivo.write(registro)
    arquivo.close()


def send_mail(mail, mail_subject, mail_message):   
    mail_from = "notificacoes@facilitasistemas.com.br"
    host = "smtplw.com.br"
    port = 587
    user = "facilitasistemas"
    password = "uqlYDTUM6733"
    try:
        s = smtplib.SMTP(host=host, port=port)
        s.starttls()
        s.login(user,password)
        message = MIMEText(mail_message, 'html')
        message['subject'] = mail_subject
        message['from'] = mail_from
        message['to'] = mail
        s.sendmail(mail_from, mail, message.as_string())
        s.quit()
        return True
    except Exception as e:
        print(e)
        s.quit()
        return False


def procurar_mensagen(cookies, cnpj, lista_email):
    url = 'https://cav.receita.fazenda.gov.br/'
    url_msg = url+'Servicos/ATSDR/CaixaPostal.app/Action/'
    with Session() as s:
        for cookie in cookies:
            s.cookies.set(cookie['name'], cookie['value'])

        pagina = s.get(url+'ecac/', verify=False)
        if pagina.status_code != 200:
            print('>>> Erro a acessar página.')
            salva_registro(cnpj, 'Erro a acessar página')
            return False
        pagina = s.get(url_msg+'ListarMensagemAction.aspx', verify=False)
        soup = BeautifulSoup(pagina.content, 'html.parser')
        tabela = soup.find('table', attrs={'id': 'gridMensagens'})

        for linha in tabela.findAll('tr'):
            colunas = linha.findAll('td')
            assunto = 'Linha de crédito criada pelo Programa Nacional de Apoio às Microempresas e Empresas de Pequeno Porte'
            texto = ''
            if colunas[4].text.strip() == assunto:
                print('>>> Encontrou mensagem de credito')
                link = url_msg+colunas[4].find('a').get('href')
                pagina = s.get(link, verify=False)
                soup = BeautifulSoup(pagina.content, 'html.parser')
                cabecalho = soup.find('div', attrs={'class': 'linhaMsgImpar'})
                corpo = soup.find('div', attrs={'class': 'txtMsg'})
                css = 'class="linhaMsgImpar" style="display:flex;width:100%;justify-content:space-between;"'
                texto = str(cabecalho).replace('class="linhaMsgImpar"', css)+str(corpo)

                for email in lista_email.split(';'):
                    if not email: continue
                    result = send_mail(email.strip(), assunto, texto)
                    #result = send_mail('rubens@veigaepostal.com.br', assunto, texto)
                    if result:
                        print(">>> Email enviado com sucesso")
                        salva_registro(cnpj, f'Email enviado com sucesso ({email.strip()})')
                    else:
                        print(">>> Erro ao enviar email")
                        salva_registro(cnpj, f'Erro ao enviar email ({email.strip()})')
                break

        if not texto:
            print('>>> Não recebeu mensagem de credito')
            salva_registro(cnpj, 'Não recebeu mensagem de credito')
        
@time_execution
def run():
    lista_auxiliar = [ # ALTERAR VARIAVEL 'result' CASO PRECISE DEFINIR UM EMAIL FIXO
        ('28885399000154', r'..\certificados\CERT RPEM 35586086.pfx', '35586086', 'rubens@veigaepostal.com.br;'),
    ]
    #lista_empresas = empresas
    lista_empresas = lista_auxiliar
    for empresa in lista_empresas[:]: 
        print('>>> Acessando empresa ', empresa[0])
        cookies, logou = login(empresa[0], empresa[1], empresa[2])
        if not logou:
            print('')
            continue
        print('>>> Feito login')
        procurar_mensagen(cookies, empresa[0], empresa[3])
        print('')

if __name__ == '__main__':
    run()
