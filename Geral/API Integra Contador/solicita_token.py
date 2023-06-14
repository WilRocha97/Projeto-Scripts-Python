# -*- coding: utf-8 -*-
import requests_pkcs12, os, requests, time, base64, json, io, chardet, pyautogui as p
from pathlib import Path
from tkinter.filedialog import askopenfilename, Tk

e_dir = Path('Execução')


def escreve_token(texto, local=e_dir, encode='latin-1'):
    if local == e_dir:
        local = Path(local)
    os.makedirs(local, exist_ok=True)

    try:
        f = open(str(local / "token.txt"), 'a', encoding=encode)
    except:
        f = open(str(local / "token - auxiliar.txt"), 'a', encoding=encode)

    f.write(texto)
    f.close()


def ask_for_file():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title='Selecione o arquivo do certificado digital',
        filetypes=[('PFX files', '*.pfx *')],
        initialdir=os.getcwd()
    )
    
    return file if file else False


# converte as credenciais para base64
def converter_base64(usuario):
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
    
    resposta = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    print(pagina.status_code)
    print(resposta)
    
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
    return resposta['jwt_token'], resposta['access_token']


def run():
    consumerKey = p.password(text='Informe a consumerKey:')
    consumerSecret = p.password(text='Informe a consumerSecret:')
    usuario = consumerKey + ":" + consumerSecret
    usuario_b64 = converter_base64(usuario)
    
    senha = p.password(text='Informe a senha do certificado digital:')
    
    # pergunta qual o arquivo do certificado
    certificado = ask_for_file()
    
    jwt_token, access_token = solicita_token(usuario_b64, certificado, senha)
    tokens = jwt_token + ' | ' + access_token
    escreve_token(tokens)
    

if __name__ == '__main__':
    run()


