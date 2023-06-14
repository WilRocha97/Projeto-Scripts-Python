# -*- coding: utf-8 -*-
import requests_pkcs12, requests, time, base64, json, io
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _escreve_header_csv, _time_execution, _open_lista_dados, _where_to_start

dados = "Dados\\Dados Integra Contador.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')
consumerKey = user[0]
consumerSecret = user[1]
usuario = consumerKey + ":" + consumerSecret

# arquivo exportado em formato PFX ou P12 com a chave privada e senha
# certificado eCNPJ ICP Brasil do contratante na loja Serpro
certificado = "Dados\\" + user[2]
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
    
    
def solicita_dctf(cnpj_contratante, cpf_certificado, cnpj_empresa, jwt_token, access_token):
    body = {
              "contratante": {
                "numero": cnpj_contratante,
                "tipo": 2
              },
              "autorPedidoDados": {
                "numero": cpf_certificado,
                "tipo": 1
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
    
    header = {'Accept': 'application/json',
               'Authorization': access_token,
               'jwt_token': jwt_token,}
    
    pagina = requests.post('https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir', data=body, headers=header)
    
    print(pagina)
    print(json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True))
    
    
def cria_pdf(pdf_base64, cnpj_empresa, nome_empresa):
    pdf_bytes = base64.b64decode(pdf_base64)
    os.makedirs('Execução/DCTFWEB', exist_ok=True)
    with open(os.path.join('Execução', 'DCTFWEB', f'DCTFWEB - {cnpj_empresa} - {nome_empresa}.pdf'), "wb") as file:
        file.write(pdf_bytes)


@_time_execution
def run():
    print(user[0], user[1], user[2], user[3])
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
        cnpj_empresa, nome_empresa = empresa
        _indice(count, total_empresas, empresa)
        jwt_token, access_token = solicita_token()
        pdf_base64 = solicita_dctf(cnpj_contratante, cpf_certificado, cnpj_empresa, jwt_token, access_token)
        cria_pdf(pdf_base64, cnpj_empresa, nome_empresa)
        

if __name__ == '__main__':
    run()


