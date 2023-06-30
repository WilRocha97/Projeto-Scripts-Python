# -*- coding: utf-8 -*-
import requests_pkcs12, os, requests, time, base64, json, io, chardet, OpenSSL.crypto, contextlib, tempfile, pyautogui as p
from pathlib import Path
from functools import wraps
from tkinter.filedialog import askopenfilename, Tk

e_dir = Path('T:\ROBO\DCTF-WEB\Execução')
e_dir_2 = Path('Execução')


def open_lista_dados(encode='latin-1'):
    file = ask_for_file_dados()
    if not file:
        return False

    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    if local == e_dir:
        local = Path(local)
    # os.makedirs(local, exist_ok=True)

    try:
        f = open(str(local / f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome}-auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


def escreve_doc(texto, nome='doc', local=e_dir, encode='latin-1'):
    if local == e_dir:
        local = Path(local)
    # os.makedirs(local, exist_ok=True)
    
    try:
        f = open(str(local / f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome} - auxiliar.txt"), 'a', encoding=encode)

    f.write(str(texto))
    f.close()


def ask_for_file(title='Selecione o Certificado Digital', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=[('PFX files', '*.pfx *')],
        initialdir=initialdir
    )
    
    return file if file else False


def ask_for_file_dados(title='Selecione a planilha de dados das empresas', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=[('Plain text files', '*.txt *.csv')],
        initialdir=initialdir
    )
    
    return file if file else False


def converter_base64(usuario):
    # converte as credenciais para base64
    return base64.b64encode(usuario.encode("utf8")).decode("utf8")


def solicita_token(usuario_b64, certificado, senha):
    headers = {
        "Authorization": "Basic " + usuario_b64,
        "role-type": "TERCEIROS",
        "content-type": "application/x-www-form-urlencoded"
    }
    body = {'grant_type': 'client_credentials'}
    pagina = requests_pkcs12.post('https://autenticacao.sapi.serpro.gov.br/authenticate',
                                  data=body,
                                  headers=headers,
                                  verify=True,
                                  pkcs12_filename=certificado,
                                  pkcs12_password=senha)
    
    resposta_string_json = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    resposta = pagina.json()
    
    try:
        escreve_doc(pagina.status_code, nome='status_code', local=e_dir_2)
        escreve_doc(resposta, nome='resposta_jason', local=e_dir_2)
        escreve_doc(resposta_string_json, nome='string_json', local=e_dir_2)
    except:
        escreve_doc(pagina.status_code, nome='status_code')
        escreve_doc(resposta, nome='resposta_jason')
        escreve_doc(resposta_string_json, nome='string_json')
    
    #
    # Output:
    #   200
    #   {
    #       "access_token": "05e658b6-212e-3f0e-b068-9d125a3ea479",
    #       "expires_in": 3032,
    #       "jwt_pucomex": null,
    #       "jwt_token": "eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIzMzY4MzExMTxj...",
    #       "scope": "am_application_scope default",
    #       "token_type": "Bearer"
    #   }
    #
    try:
        return resposta['jwt_token'], resposta['access_token']
    except:
        return resposta['message'], resposta['description']
    

def solicita_dctf(cnpj_contratante, cnpj_empresa, access_token):
    comp = p.prompt(text='Informe a competência das guias que deseja solicitar', default='00/0000')
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]
    
    body = {
              "contratante": {
                "numero": cnpj_contratante,
                "tipo": 2
              },
              "autorPedidoDados": {
                "numero": cnpj_contratante,
                "tipo": 2
              },
              "contribuinte": {
                "numero": cnpj_empresa,
                "tipo": 2
              },
              "pedidoDados": {
                "idSistema": "DCTFWEB",
                "idServico": "GERARGUIA31",
                "versaoSistema": "1.0",
                "dados": "{\"categoria\": \"GERAL_MENSAL\",\"anoPA\":\"" + ano + "\",\"mesPA\":\"" + mes + "\"}"
              }
            }
    
    header = {'Content-Type': 'application/json',
               'Authorization': access_token}
    
    pagina = requests.post('https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir', json=body, headers=header)
    
    resposta = pagina.json()
    resposta_string_json = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    
    try:
        escreve_doc(resposta, nome='resposta_jason_guia', local=e_dir_2)
        escreve_doc(resposta_string_json, nome='string_json_guia', local=e_dir_2)
        escreve_doc(resposta['dados'], nome='pdf_base_64', local=e_dir_2)
    except:
        escreve_doc(resposta, nome='resposta_jason_guia')
        escreve_doc(resposta_string_json, nome='string_json_guia')
        escreve_doc(resposta['dados'], nome='pdf_base_64')
        
    return mes, ano, resposta['dados']['PDFByteArrayBase64'], resposta['mensagens'][0]['texto']

    
def cria_pdf(pdf_base64, cnpj_empresa, nome_empresa, mes, ano):
    pdf_bytes = base64.b64decode(pdf_base64)
    # os.makedirs('Execução/DCTFWEB', exist_ok=True)
    with open(os.path.join('Execução', 'Guias', f'DCTFWEB {mes}-{ano} - {cnpj_empresa} - {nome_empresa}.pdf'), "wb") as file:
        file.write(pdf_bytes)


def run():
    cnpj_contratante = p.prompt(text='Informe o CNPJ do contratante do serviço SERPRO')
    consumerKey = p.password(text='Informe a consumerKey:')
    consumerSecret = p.password(text='Informe a consumerSecret:')
    
    usuario = consumerKey + ":" + consumerSecret
    usuario_b64 = converter_base64(usuario)
    
    senha = p.password(text='Informe a senha do certificado digital:')
    
    # pergunta qual o arquivo do certificado
    certificado = ask_for_file()
    
    jwt_token, access_token = solicita_token(usuario_b64, certificado, senha)
    
    tokens = jwt_token + ' | ' + access_token
    try:
        escreve_doc(tokens, nome='tokens', local=e_dir_2)
    except:
        escreve_doc(tokens, nome='tokens')
        
    # abrir a planilha de dados
    empresas = open_lista_dados()
    if not empresas:
        return False

    for count, empresa in enumerate(empresas, start=1):
        cnpj_empresa, nome_empresa = empresa
        mes, ano, pdf_base64, mensagens = solicita_dctf(cnpj_contratante, cnpj_empresa, access_token)
        try:
            cria_pdf(pdf_base64, cnpj_empresa, nome_empresa, mes, ano)
            mensagen_2 = ''
        except Exception as e:
            mensagen_2 = f'Não gerou PDF {e}'
        
        try:
            escreve_relatorio_csv(f'{cnpj_empresa};{nome_empresa};{mensagens};{mensagen_2}', local=e_dir_2)
        except:
            escreve_relatorio_csv(f'{cnpj_empresa};{nome_empresa};{mensagens};{mensagen_2}')
        
    p.alert(text='Robô finalizado!')
    
    
if __name__ == '__main__':
    run()


