from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from requests import Session
from pandas import read_excel
import warnings
import base64
import json
import time
import os
import sys
sys.path.append("..")
from comum import escreve_relatorio_csv, time_execution


warnings.filterwarnings('ignore')

_SITUACOES = [
    'Atenção:Código de Acesso expirado.',
    'Atenção:Dados Inválidos. Favor conferir senha e código digitados.'
]


def wait_to_load(driver):
    for i in range(60):
        time.sleep(1)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        not_loading = soup.find('div', attrs={'class':'notLoading'})
        if not_loading:
            time.sleep(2)
            return True
    else:
        return False


def download_irpf(nome, cpf, acesso, senha, quota, venc):  
    url_login = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    url_main = 'https://cav.receita.fazenda.gov.br/ecac/'
    url_irpf = 'https://www3.cav.receita.fazenda.gov.br/extratodirpf/#/'
    url_debito = "https://www3.cav.receita.fazenda.gov.br/extratodirpf/#/debitos/2021"
    url_extrato = "https://www3.cav.receita.fazenda.gov.br/extratodirpf/api/extrato/demonstrarDebitos/2021"
    url_darf = "https://www3.cav.receita.fazenda.gov.br/extratodirpf/api/extrato/gerarDarf/"
    nome = nome.strip()

    if not all([nome, cpf, acesso, senha]):
        print(f">>>{nome} - Algum dos dados de acesso estava em branco")
        escreve_relatorio_csv(f"{nome};Atenção:Dados Inválidos. Favor conferir senha e código digitados.")
        return True

    for i in range(5):
        try:
            print("Tentando Logar", cpf)
            driver = webdriver.PhantomJS('phantomjs.exe')
            driver.set_page_load_timeout(30)
            driver.set_window_size(1366,768)
            
            driver.get(url_login)

            driver.execute_script(f"document.querySelector('#NI').value = '{cpf}';document.querySelector('#CodigoAcesso').value = '{acesso}';document.querySelector('#Senha').value = '{senha}'")
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="login-dados-usuario"]/p[4]/input').click()
            time.sleep(2)

            if driver.find_elements_by_class_name("tipsy-inner"):
                print(f">>>{nome} - Algum dos dados de acesso estava errado")
                escreve_relatorio_csv(f"{nome};Atenção:Dados Inválidos. Favor conferir senha e código digitados.")
                driver.quit()
                return True
            elif driver.current_url != url_main:
                soup =  BeautifulSoup(driver.page_source, 'html.parser')
                div_validation = soup.find('div', attrs={'class':'login-caixa-erros-validacao'})
                print(f">>>{nome} - {div_validation.find('p').text}")
                escreve_relatorio_csv(f"{nome};{div_validation.find('p').text}")
                driver.quit()
                return True
            break
        except:
            print("\n>>>Nova tentativa: login")
            driver.quit()
            time.sleep(5)
    else:
        print(f">>>{nome} um erro ocorreu, tentativas excedidas 1")
        driver.quit()
        return False

    for i in range(5):
        print("Verificando quadro servicos")
        try:
           driver.get(url_irpf)
        except:
           print("\n>>>Nova tentativa: timeout")
           continue
        time.sleep(2)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        quadro_servicos = soup.find('div',attrs={'id':'tServicos'})
        if quadro_servicos: break
        print("\n>>>Nova tentativa: sem quadro")
    else:
        print(f">>>{nome} um erro ocorreu, tentativas excedidas 2")
        driver.quit()
        return False

    if not wait_to_load(driver):
        print(f">>>{nome} um erro ocorreu, tentativas excedidas 3")
        driver.quit()
        return False
        
    try: #fechar modal que as vezes aparece
        driver.execute_script("document.querySelector('#helppanel_text > div.campo_sele_help > div.btClos.glyphicon.glyphicon-remove').click()")
    except:
        pass

    driver.get(url_debito)
    time.sleep(2)

    session = Session()
    if driver.current_url == url_debito:
        for cookie in driver.get_cookies():
            session.cookies.set(cookie['name'], cookie['value'])
    else:
        print(">>> Falha durante acesso as guias.")
        driver.quit()
        session.close()

    info = driver.execute_script("return sessionStorage.getItem('_extirpf');")
    token = json.loads(info)
    token = json.loads(token['token'])
    header = {
        'Authorization': token['type']+" "+token['key'],
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    driver.quit()

    pagina = session.get(url_extrato, headers=header)
    info = json.loads(pagina.text)
    debitos = info["debitosIAP"][0]
    numDebito = debitos['numeroDebito']
    cotas = debitos['cotas']
    for cota in cotas:
        if cota['numeroSequencial'] == quota and cota['situacaoCotaFmt'] == 'A vencer':
        #if cota['dataVencimento'] == venc and cota['situacaoCotaFmt'] == 'A vencer':
            str_dados = f"{numDebito}/{quota}/{venc}"
            pagina = session.get(url_darf+str_dados, headers=header)
            dados = json.loads(pagina.text)
            # Pega os dados da guia e converte novamente para bytes
            base = dados['darfPdf'].encode('ascii')
            dados_bytes = base64.b64decode(base)
            vencimento = cota['dataVencimentoFmt'].replace('/', '-')
            arq = f"DARF {vencimento} {nome}.pdf"

            try:
                arquivo = open(os.path.join('documentos', arq), 'wb')
            except FileNotFoundError:
                os.makedirs('documentos', exist_ok=True)
                arquivo = open(os.path.join('documentos', arq), 'wb')
            arquivo.write(dados_bytes)
            arquivo.close()
            print(f">>> Arquivo {arq} salvo")
            escreve_relatorio_csv(f"{nome};Parcela {quota} encontrada")
            break
    else:
        escreve_relatorio_csv(f"{nome};Sem parcelas em aberto")
        print('>>> Sem parcelas em aberto')

    session.close()
    return True


def analisa_planilha(planilha):
    lista = read_excel(planilha, sheet_name="LISTA GERAL", keep_default_na=False, header=0)
    lista.query("QUOTAS != ''", inplace=True)
    lista_empresas = [[x.NOME, x.CPF, x.CÓDIGO, x.SENHA, x.QUOTAS] for x in lista.itertuples()]
    
    return lista_empresas


@time_execution
def run():
    _PLANILHA = "IRPF 2020 - QUOTAS.xls.xlsx"
    vencimento = '20210129' # ultimo dia util do mês (yyyymmdd)

    lista_empresas = analisa_planilha(_PLANILHA)
    total = len(lista_empresas)
    for index, dados in enumerate(lista_empresas, 1):
        try:
            download_irpf(*dados, vencimento)
        except Exception as e:
            print('>>>> Ocorreu um erro')
            print(e)
        print(f"\n--- Finalizada empresa {index}/{total} ---")
        print("\n--- Proxima empresa ---")


if __name__ == "__main__":
    run()