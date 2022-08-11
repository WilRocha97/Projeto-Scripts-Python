from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from Dados import empresas
import time
import os
import sys
sys.path.append("..")
from comum import escreve_relatorio_csv, time_execution



_SITUACOES = [
    'Atenção:Código de Acesso expirado.',
    'Atenção:Dados Inválidos. Favor conferir senha e código digitados.'
]


def salva_sem_senha():
    arquivo = open('Resumo.csv', 'r')
    lista_empresas = arquivo.readlines()
    arquivo.close()
    new = open('Empresas sem senha.csv', 'w')
    for linha in lista_empresas:
        empresa, situacao = linha.split(';')
        if situacao.strip() in _SITUACOES:
            new.write(linha)

    new.close()


def salva_excedidos(excedidos):
    try:
        arquivo = open("Excedidos.py","a")
    except:    
        arquivo = open("Excedidos.py","w")

    lista_dados = ['"'+dado+'"' for dado in excedidos]
    arquivo.write('('+','.join(lista_dados)+'),\n')
    arquivo.close()


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


def print_pendencias(driver, nome, cpf, i_alerta, alerta):
    anos = []
    children = alerta.findChildren('a', attrs={'class':'badge link'})
    for i_child, child in enumerate(children, 2): 
        ano = child.text
        anos.append(ano)
        driver.execute_script(f"document.querySelector('#tAlertas > ul > li:nth-child({i_alerta}) > div > span:nth-child({i_child}) > a').click()")
        if not wait_to_load(driver): break
        driver.save_screenshot(f"./relatorios/pendencia/{nome} {cpf} - pendencia_{ano}.png")
        driver.execute_script(f"document.querySelector('#menu_item_{ano}').children[0].click()")
        if not wait_to_load(driver): break
        driver.save_screenshot(f"./relatorios/pendencia malha/{nome} {cpf} - pendencia_malha_{ano}.png")
        driver.execute_script("document.querySelector('#upperBar > section > div.counBts.ng-star-inserted > div.home > a').click();")   
        if not wait_to_load(driver): break
    else:
        return (True, anos)
    return (False, anos)
    

def consulta_irpf(nome, cpf, acesso, senha):  
    url_login = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    url_main = 'https://cav.receita.fazenda.gov.br/ecac/'
    url_irpf = 'https://www3.cav.receita.fazenda.gov.br/extratodirpf/#/'

    if not all([nome, cpf, acesso, senha]):
        print(f">>>{nome} - Algum dos dados de acesso estava em branco")
        escreve_relatorio_csv(f"{nome};Atenção:Dados Inválidos. Favor conferir senha e código digitados.")
        return True

    for i in range(5):
        try:
            print("Tentando Logar")
            driver = webdriver.PhantomJS(os.path.join('..', 'phantomjs.exe'))
            driver.set_page_load_timeout(30)
            driver.set_window_size(1366,768)
            
            driver.get(url_login)

            driver.execute_script(f"document.querySelector('#NI').value = '{cpf}';document.querySelector('#CodigoAcesso').value = '{acesso}';document.querySelector('#Senha').value = '{senha}'")
            time.sleep(1)
            #driver.execute_script("document.querySelector('#login-dados-usuario > p:nth-child(4) > input').click()")
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
            pass
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
        time.sleep(5)

        soup = BeautifulSoup(driver.page_source, 'html.parser')
        quadro_servicos = soup.find('div',attrs={'id':'tServicos'})
        if quadro_servicos: break

        driver.save_screenshot(f"./relatorios/excecoes/{nome}{i}.png")
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

    driver.save_screenshot(f"./relatorios/geral/{nome} {cpf} - geral.png")
    print("print geral")

    quadro_alerta = BeautifulSoup(driver.page_source, 'html.parser').find('div',attrs={'id':'tAlertas'})
    if quadro_alerta:
        print("Verificando pendencias")
        result = False
        anos_pendentes = []
        alertas = quadro_alerta.findAll('div')
        for i_alerta, alerta in enumerate(alertas, 1):
            excecoes = [
                "Não existem dispositivos móveis (celular / tablet) habilitados para acompanhar as informações relativas a declaração de IRPF. Para maior comodidade e agilidade, realize o cadastramento.",
                "O débito da quota do IRPF foi agendada junto ao banco. Acompanhe o extrato bancário no vencimento para verificar se o débito foi efetuado.",
                "A declaração está na fila de restituição. Acompanhe os próximos lotes.",
                "Contribuinte consta como dependente em outra declaração."
            ]

            if alerta.get('title') not in excecoes:
                result, anos = print_pendencias(driver, nome, cpf, i_alerta, alerta)
                if not result:
                    print(f">>>{nome} um erro ocorreu, tentativas excedidas 4")
                    driver.quit()
                    return False
                anos_pendentes += anos
        else:
            if result:
                str_anos = ", ".join(anos_pendentes)
                print(f">>>{nome} empresa com pendencias nos anos {str_anos}")
                escreve_relatorio_csv(f"{nome};empresa com pendencia nos anos {str_anos}")    
            else:
                print(f">>>{nome} empresa sem pendencias")
                escreve_relatorio_csv(f"{nome};empresa sem pendencias")
    else:
        print(f">>>{nome} empresa sem pendencias")
        escreve_relatorio_csv(f"{nome};empresa sem pendencias")

    driver.quit()
    return True


@time_execution
def run():
    os.makedirs('relatorios/geral', exist_ok=True)
    os.makedirs('relatorios/excecoes', exist_ok=True)
    os.makedirs('relatorios/pendencia', exist_ok=True)
    os.makedirs('relatorios/pendencia malha', exist_ok=True)

    empresas_auxiliar = [
    ]
    lista_empresas = empresas[:]
    #lista_empresas = empresas_auxiliar
    total = len(lista_empresas)
    for index, dados in enumerate(lista_empresas, 1):
        #if dados[1] not in buscar: continue
        try:
            result = consulta_irpf(*dados)
            if not result:
                #lista_empresas.append(*dados)
                salva_excedidos(dados)
        except Exception as e:
            print(e)
        print(f"\n--- Finalizada empresa {index}/{total} ---")
        print("\n--- Proxima empresa ---")
    salva_sem_senha()

if __name__ == "__main__":
    run()