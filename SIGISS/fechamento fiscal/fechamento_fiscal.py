# -*- coding: utf-8 -*-
# ! python3
from requests import Session
from datetime import datetime
import re


def consulta_livro_fiscal(cnpj, dtInicio, dtFinal, ano, fechar):
    url_acesso = "https://valinhos.sigissweb.com/ControleDeAcesso"
    url_emite = "https://valinhos.sigissweb.com/nfecentral?oper=listarnfe"
    url_emitiu = "https://valinhos.sigissweb.com/nfecentral?oper=efetuapesquisasimples"
    url_fechou = "https://valinhos.sigissweb.com/lancamentocentral?oper=efetuapesquisasimples"
    url_fechar = "https://valinhos.sigissweb.com/lancamentocentral?"
    # obter login e senha do banco de dados utilizando o cnpj
    login = cnpj
    senha = "1217800"

    # inicia a sessão no site, realiza o login e obtem os arquivos para o contribuinte
    for x in range(2):
        with Session() as s:
            print("login:" + str(login))
            login_data = {"loginacesso":login,"senha":senha}
            pagina = s.post(url_acesso, login_data)
            str_logou = "Login e/ou Senha inv"
            nao_logou = re.search(str_logou, pagina.text)
            print(nao_logou)

            if pagina.status_code != 200:
                print('>>> Erro a acessar página.')
                return False

            #Confere se emite NF
            str_emite = "não possui nenhuma solicitação para emissão de NFe aprovada."
            response_emite = s.get(url_emite)
            result_emite = re.search(str_emite, response_emite.text)

            if result_emite is None:
                print(">>> Empresa emite")

                #Confere se emitiu NF
                for competencia in range(int(dtInicio), int(dtFinal)+1):
                    str_competencia = str(competencia).rjust(2,"0")
                    info_emitiu = {
                        'cnpj_cpf_destinatario': '', 'operCNPJCPFdest': 'EX', 'RAZAO_SOCIAL_DESTINATARIO': '',
                        'selnomedestoper': 'EX', 'id_codigo_servico': '', 'serie': '', 'numero_nf': '',
                        'operNFE': '=', 'numero_nf2': '', 'rps': '', 'operRPS': '=', 'rps2': '',
                        'data_emissao': '', 'operData': '=', 'data_emissao2': '', 'mesi': str_competencia,
                        'anoi': ano, 'mesf': str_competencia, 'anof': ano, 'aliq_iss': '',
                        'regime': '?', 'iss_retido': '?', 'cancelada': '?', 'tipoPesq': 'normal'
                    }

                    str_emitiu = "Não há Dados a serem exibidos para a Pesquisa efetuada"
                    response_emitiu = s.post(url_emitiu, info_emitiu)

                    if response_emitiu.status_code != 200:
                        print('>>> Erro a acessar página.')
                        return False

                    result_emitiu = re.search(str_emitiu, response_emitiu.text)

                    if result_emitiu is not None:
                        print(">>> Não teve movimento")

                        #Confere se o livro fiscal foi fechado
                        info_fechou = {
                            'tomador_prestador': 'P', 'cnpj_cpf_decl': '', 'selcnpjcpfoperdecl': 'EX',
                            'cnpj_cpf_dest': '', 'selcnpjcpfoperdest': 'EX', 'cidade_dest': '',
                            'selcidadedestoperdest': 'EX', 'mesi': str_competencia, 'anoi': ano,
                            'mesf': str_competencia, 'anof': ano, 'mesicaixa': '', 'anoicaixa': '',
                            'mesfcaixa': '', 'anofcaixa': '', 'iss_retido_fonte': '?', 'regime': '?',
                            'movimento': '?', 'documento_fiscal_canc': '?', 'num_docu_fiscal_pesq': '',
                            'serie_docu_fiscal': '', 'data': '', 'aliquota': '?', 'caixaindicado': '?',
                            'classif': '?'
                        }

                        str_fechou = "Não há dados a serem exibidos para a pesquisa efetuada"
                        response_fechou = s.post(url_fechou, info_fechou)
                        if response_fechou.status_code != 200:
                            print('>>> Erro a acessar página.')
                            return False

                        result_fechou = re.search(str_fechou, response_fechou.text)
                        if result_fechou is not None:
                            print(">>> Não fechou o livro fiscal")

                            #Fechar o livro fiscal se flag True
                            if fechar:
                                info_fechar = {
                                    'mes': str_competencia, 'ano': ano, 'tomador_prestador': 'P',
                                    '': '', 'movimento': 'N', '': '', 'regime': 'S', '': '', '': '',
                                    '': '', '': '', '': '', 'cnpj_cpf_decl': str(cnpj),
                                    str(cnpj): str(cnpj), 'regimeempresa': 'S', 'edomunicipio': 'true',
                                    'cnpj_cpf_dest': '', 'classif': '', 'documento_fiscal_canc': 'N',
                                    'iss_retido_fonte': 'N', 'num_docu_fiscal': '', 'serie_docu_fiscal': '',
                                    'data': '', 'valor_docu_fiscal': 'R$ 0,00', 'deducoes_legais': 'R$ 0,00',
                                    'valor_servicos': 'R$ 0,00', 'aliquota': '', 'valor_imposto': 'R$ 0,00',
                                    'id_lancamento': '0', 'operacao': 'inssalvar', 'id_dest': '0',
                                    'somenteLeitura': 'false', 'naopodealteraraliquota': 'false', 
                                    'tipoinclusao': 'M', 'btnsalvar': 'Salvar', 'btncancelar': 'Cancelar',
                                    '': '', 'oper': 'consistirEGravar', 'tipoInsert': 'M'
                                }
                                #response_fechar = s.post(url_fechar, info_fechar)
                                #if response_fechar.status_code != 200:
                                #    print('>>> Erro a acessar página.')
                                #    return False   
                                #print(">>> Livro fechado")
                        else:
                            print(">>> Fechou o livro fiscal")
                    else:
                        print(">>> Teve movimento")
            else:
                print(">>> Empresa não emite")

        #login = '24982859000101'
        senha = '2376500'
    return True


if __name__ == '__main__':
    comeco = datetime.now()
    cnpj='06922286000149' 
    dtInicio='01'
    dtFinal='01'
    ano='2020'
    fechar = True
    concluido = consulta_livro_fiscal(cnpj, dtInicio, dtFinal, ano, fechar)
    if concluido:
        print('\n>>> Programa concluido com sucesso.')
    else:
        print('\n>>> Falha durante execução do programa.')
    fim = datetime.now()
    print("\n>>> Tempo total:", (fim-comeco))
    