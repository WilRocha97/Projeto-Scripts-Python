# -*- coding: utf-8 -*-
import requests_pkcs12, requests, time, base64, json
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _escreve_header_csv, _time_execution, _open_lista_dados, _where_to_start

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Integra Contador.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')
consumerKey = user[0]
consumerSecret = user[1]
usuario = consumerKey + ":" + consumerSecret

# arquivo exportado em formato PFX ou P12 com a chave privada e senha
# certificado eCNPJ ICP Brasil do contratante na loja Serpro
certificado = "V:\\Setor Robô\\Scripts Python\\Geral\\API Integra Contador\\ignore\\" + user[2]
senha = user[3]

# converte as credenciais para base64
def converter_base64():
    return base64.b64encode(usuario.encode("utf8")).decode("utf8")
    

def token():
    usuario_b64 = converter_base64()
    
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
    
    print(pagina.status_code)
    print(json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True))
    time.sleep(22)
    
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
    

@_time_execution
def run():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, numero = empresa
        _indice(count, total_empresas, empresa)
        token = solicita_token()
        

if __name__ == '__main__':
    run()


