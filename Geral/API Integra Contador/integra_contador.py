# -*- coding: utf-8 -*-
import requests_pkcs12, os, requests, time, base64, json, io, chardet, OpenSSL.crypto, contextlib, tempfile, pyautogui as p
from pathlib import Path
from functools import wraps
from tkinter.filedialog import askopenfilename, Tk

e_dir = Path('T:\ROBO\DCTF-WEB\Execução')
e_dir_api = Path('T:\ROBO\DCTF-WEB\Execução\Response')
e_dir_2 = Path('Execução')


# abre a lista de dados da empresa em .csv
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


# escreve os andamentos das requisições em um .csv
def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(str(local / f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome}-auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()


# escreve arquivos .txt com as respostas da API
def escreve_doc(texto, nome='doc', local=e_dir_api, encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(str(local / f"{nome}.txt"), 'a', encoding=encode)
    except:
        f = open(str(local / f"{nome} - auxiliar.txt"), 'a', encoding=encode)

    f.write(str(texto))
    f.close()


# abre o arquivo .pfx do certificado digital
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


# seleciona o arquivo .csv com as dados das empresas
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


# converte as chaves de acesso do site da API em base64 para fazer a requisição das chaves de acesso para o serviço
def converter_base64(usuario):
    # converte as credenciais para base64
    return base64.b64encode(usuario.encode("utf8")).decode("utf8")


# solicita as chaves de acesso para o serviço
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
    
    # anota as respostas para tratar possíveis erros
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
    

# solicita a guia de DCTF WEB na API
def solicita_dctf(comp, cnpj_contratante, id_empresa, access_token, jwt_token):
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]
    
    # verifica se a guia será para CPF ou CNPJ, se for CPF é código 1 e CNPJ é código 2
    if len(id_empresa) > 12:
        cod_consulta = '2'
    else:
        cod_consulta = '1'
    
    data = {
              "contratante": {
                "numero": str(cnpj_contratante),
                "tipo": 2
              },
              "autorPedidoDados": {
                "numero": str(cnpj_contratante),
                "tipo": 2
              },
              "contribuinte": {
                "numero": str(id_empresa),
                "tipo": int(cod_consulta)
              },
              "pedidoDados": {
                "idSistema": "DCTFWEB",
                "idServico": "GERARGUIA31",
                "versaoSistema": "1.0",
                "dados": "{\"categoria\": \"GERAL_MENSAL\",\"anoPA\":\"" + str(ano) + "\",\"mesPA\":\"" + str(mes) + "\"}"
              }
            }
    
    headers = {'accept': 'text/plain',
               'Authorization': 'Bearer ' + access_token,
               'Content-Type': 'application/json',
               'jwt_token': jwt_token}
    
    pagina = requests.post('https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir', headers=headers, data=json.dumps(data))
    resposta = pagina.json()
    resposta_string_json = json.dumps(json.loads(pagina.content.decode("utf-8")), indent=4, separators=(',', ': '), sort_keys=True)
    
    # anota as respostas da API para tratar possíveis erros
    try:
        os.makedirs(e_dir, exist_ok=True)
        escreve_doc(resposta, nome='resposta_jason_guia', local=e_dir_2)
        escreve_doc(resposta_string_json, nome='string_json_guia', local=e_dir_2)
        escreve_doc(resposta['dados'], nome='pdf_base_64', local=e_dir_2)
    except:
        escreve_doc(resposta, nome='resposta_jason_guia')
        escreve_doc(f'{id_empresa}\n{resposta_string_json}', nome='string_json_guia')
        escreve_doc(resposta['dados'], nome='pdf_base_64')
    #
    # output
    # {
    #     "contratante":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "autorPedidoDados":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "contribuinte":
    #         {
    #             "numero": "00000000000000",
    #             "tipo": 2
    #         },
    #     "pedidoDados":
    #         {
    #             "idSistema": "DCTFWEB",
    #             "idServico": "GERARGUIA31",
    #             "versaoSistema": "1.0",
    #             "dados": "{\"categoria\": \"GERAL_MENSAL\",\"anoPA\":\"2027\",\"mesPA\":\"11\"}"
    #         },
    #     "status": 200,
    #     "dados": "{\"PDFByteArrayBase64\":\"JVBERi0xLjUKJafj8fEKMi..."}",
    #     "mensagens":
    #         [
    #             {
    #                 "codigo": "Aviso-DCTFWEB-MG11",
    #                 "texto": "Emissor Guia Pagamento executado com sucesso."
    #             },
    #             {
    #                 "codigo": "Sucesso-DCTFWEB-MG00",
    #                 "texto": "Requisição efetuada com sucesso"
    #             }
    #         ]
    # }
    #
    try:
        dados_pdf = json.loads(resposta["dados"])
        return mes, ano, dados_pdf["PDFByteArrayBase64"], resposta['mensagens'][0]['texto']
    except:
        try:
            return mes, ano, resposta['dados'], resposta['mensagens'][2]['texto']
        except:
            return mes, ano, resposta['dados'], resposta['mensagens'][0]['texto']
        

# cria o PDF usando os bytes retornados da requisição na API
def cria_pdf(pdf_base64, id_empresa, nome_empresa, mes, ano):
    # limpa o nome da empresa para não dar erro no arquivo
    nome_empresa = nome_empresa.replace('/', ' ').replace(',', '')
    
    # verifica se a pasta para salvar o PDF existe, se não então cria
    e_dir_guias = Path('T:\ROBO\DCTF-WEB\Execução\Guias ' + mes + '-' + ano)
    os.makedirs(e_dir_guias, exist_ok=True)
    
    # decodifica a base64 em bytes
    pdf_bytes = base64.b64decode(pdf_base64)
    # cria o PDF a partir dos bytes
    with open(os.path.join(e_dir_guias, f'DCTFWEB {mes}-{ano} - {id_empresa} - {nome_empresa}.pdf'), "wb") as file:
        file.write(pdf_bytes)


def run():
    # pega o CNPJ da empresa contratante da API
    cnpj_contratante = p.prompt(text='Informe o CNPJ do contratante do serviço SERPRO')
    # pega as chaves de acesso encontradas no site da API
    consumer_key = p.password(text='Informe a consumerKey:')
    consumer_secret = p.password(text='Informe a consumerSecret:')
    # concatena os tokens para gerar um único em base64
    usuario = consumer_key + ":" + consumer_secret
    usuario_b64 = converter_base64(usuario)
    
    # pega a senha do certificado digital usado para cadastrar no site da API
    senha = p.password(text='Informe a senha do certificado digital:')
    
    # pergunta qual o arquivo do certificado
    certificado = ask_for_file()
    
    # limpa a pasta de response da api
    try:
        for arquivo in os.listdir(e_dir_api):
            os.remove(os.path.join(e_dir_api, arquivo))
    except:
        pass
    
    # solicita os tokens para realizar a emissão das guias
    jwt_token, access_token = solicita_token(usuario_b64, certificado, senha)
    
    tokens = jwt_token + ' | ' + access_token
    try:
        os.makedirs(e_dir, exist_ok=True)
        escreve_doc(tokens, nome='tokens', local=e_dir_2)
    except:
        escreve_doc(tokens, nome='tokens')
        
    # abrir a planilha de dados
    empresas = open_lista_dados()
    if not empresas:
        return False
    
    # pega a competência das guias que serão emitidas
    comp = p.prompt(text='Informe a competência das guias que deseja solicitar', default='00/0000')
    for count, empresa in enumerate(empresas, start=1):
        id_empresa, nome_empresa = empresa
        
        # solicita a guia de DCTF
        mes, ano, pdf_base64, mensagens = solicita_dctf(comp, cnpj_contratante, id_empresa, str(access_token), str(jwt_token))
        
        # se não retornar o PDF não precisa da segunda mensagem
        if not pdf_base64:
            mensagen_2 = ''
        # se retornar o PDF
        else:
            try:
                # tenta converter a base64 em PDF e não precisa da segunda mensagem
                cria_pdf(pdf_base64, id_empresa, nome_empresa, mes, ano)
                mensagen_2 = ''
            # se não converter o PDF captura a segunda mensagem
            except Exception as e:
                mensagen_2 = f'Não gerou PDF {e}'
        
        try:
            # tenta escrever os andamentos em um local, se não conseguir, escreve no outro
            os.makedirs(e_dir, exist_ok=True)
            escreve_relatorio_csv(f'{id_empresa};{nome_empresa};{mensagens};{mensagen_2}', nome=f'Andamentos DCTF-WEB {mes}-{ano}', local=e_dir_2)
        except:
            escreve_relatorio_csv(f'{id_empresa};{nome_empresa};{mensagens};{mensagen_2}', nome=f'Andamentos DCTF-WEB {mes}-{ano}')
        
    p.alert(text='Robô finalizado!')
    
    
if __name__ == '__main__':
    run()
