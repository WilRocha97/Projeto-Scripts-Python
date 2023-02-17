import requests
import pandas as pd


def escreve_relatorio_csv(texto):
    try:
        arquivo = open("Resumo.csv", 'a')
    except:
        arquivo = open("Resumo.csv", 'w')
    arquivo.write(texto+'\n')
    arquivo.close()


def main():
    urls = {}
    h = ['NOME', 'CPF', 'dt']
    escreve_relatorio_csv('CPF;SITUACAO_RETORNO;CPF_RETORNO;NOME_RETORNO;DATA_NASCIMENTO_RETONO;SITUACAO_CADASTRAL_RETORNO;DATA_INSCRICAO_RETORNO;DIGITO_VERIFICADOR_RETORNO;COMPROVANTE_EMITIDO_RETORNO;COMPROVANTE_EMITIDO_DATA_RETORNO')
    arquivo = pd.read_excel('pesquisa.xlsx', names=h, header=0)
    #filtro = arquivo.query("dt != 0").drop_duplicates(subset=['dt', 'CPF'])
    filtro = arquivo.drop_duplicates(subset=['dt', 'CPF'])
    for item in filtro.itertuples():
        try:
            data = item.dt.strftime('%d/%m/%Y')
        except:
            data = item.dt
        cpf = ''.join([x for x in str(item.CPF) if x.isdigit()])
        cpf = cpf.rjust(11, '0')
        url = f'http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf={cpf}&data={data}&token=8069395DeIFwyYJkP14569048&ignore_db'
        urls[cpf] = url

    #urls = {
    #    '22010251857': 'http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf=22010251857&data=19/05/1971&token=97690560gAzGdjWLCa176377344',
    #    '34437394828': 'http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf=34437394828&data=27/05/2010&token=97690560gAzGdjWLCa176377344',
    #    '09899251887': 'http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf=09899251887&data=07/08/1947&token=97690560gAzGdjWLCa176377344',
    #    '21982958812': 'http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf=21982958812&data=02/12/2011&token=97690560gAzGdjWLCa176377344',
    #}
    for cpf, url in urls.items():
        print('Analisando', cpf)
        while True:
            try:
                data = requests.get(url)
            except:
                print('Erro durante consulta.')
            else:
                break
        data = data.json()
        if data['return'] == 'NOK':
            escreve_relatorio_csv(f"{cpf};{data.get('message', 'Dados invalidos')}")
        else:
            result = data['result']
            itens = [
                cpf, data['return'], result['numero_de_cpf'],
                result['nome_da_pf'], result['data_nascimento'],
                result['situacao_cadastral'], result['data_inscricao'],
                result['digito_verificador'], result['comprovante_emitido'],
                result['comprovante_emitido_data']
            ]

            escreve_relatorio_csv(';'.join(itens))


if __name__ == '__main__':
    main()
