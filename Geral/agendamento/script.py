# -*- coding: utf-8 -*-
from selenium.webdriver.support.ui import Select
from urllib3 import disable_warnings, exceptions
from datetime import datetime, timedelta
from selenium import webdriver
from requests import Session
from threading import Thread
from time import sleep
from json import loads
import sys

disable_warnings(exceptions.InsecureRequestWarning)
# para ignorar os alertas os erros gerados pelo certificado do site que esta vencido


def data_valida():
    data = datetime.now()
    semana = data.weekday()
    if semana == 4:
        data += timedelta(days=3)
    else:
        data += timedelta(days=1)
    data = data.strftime("%d/%m/%Y")

    # DATA FIXA PARA POSSÍVEIS FERIADOS
    # data = datetime(2021, 1, 4).strftime("%d/%m/%Y")

    return data


def agendar_horario(login, senha, data_conferida):
    url = "https://senhafacil.com.br/agendamento/"
    url_busca = "https://senhafacil.com.br/api/v1/configuracao/agenda/buscarAgendaData"
    url_agenda = "https://senhafacil.com.br/api/v1/agendamento/agendar"
    header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    # options.add_argument("--start-maximized") #Essa opção não funciona com multiplas thread
    options.add_argument("--log-level=3")
    options.add_argument("--ignore-certificate-errors")
    driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)

    driver.get(url)
    sleep(2)
    # REALIZA O LOGIN NO SITE
    driver.find_element_by_xpath("/html/body/div[2]/header/div/ul/li[1]/button").click()

    sleep(2)
    for i in range(5):
        try:
            cpf = driver.find_element_by_xpath('//*[@id="login-general"]/form/div[1]/div[4]/div/label/input')
            cpf.click()
            break
        except Exception as e:
            print("\t", e)
            sleep(0.5)
    else:
        raise Exception("Tentativas excedidas")
    cpf.send_keys(login)
    password = driver.find_element_by_id("senha")
    password.click()
    password.send_keys(senha)
    driver.find_element_by_xpath('//*[@id="login-general"]/form/div[3]/div/button').click()
    sleep(3)

    # SELECIONA A CIDADE
    select = Select(driver.find_element_by_xpath('//*[@id="main-banner"]/div[2]/form/div[1]/select'))
    select.select_by_visible_text('JUNDIAÍ')
    # select.select_by_visible_text('JALES')
    sleep(1)

    select = Select(driver.find_element_by_xpath('//*[@id="main-banner"]/div[2]/form/div[2]/select'))
    select.select_by_visible_text('ATENDIMENTO')

    select = Select(driver.find_element_by_xpath('//*[@id="main-banner"]/div[2]/form/div[3]/div[1]/select'))
    select.select_by_visible_text('ATENDIMENTO REMOTO')

    driver.find_element_by_xpath('//*[@id="main-banner"]/div[2]/form/button[2]').click()
    sleep(4)

    # ACEITA OS TERMOS
    driver.find_element_by_id('aceiteDocsObrigatoriosServico').click()
    driver.find_element_by_xpath('//*[@id="form-aceite_docs_obrigatorios-servico"]/form/div[3]/div[2]/button').click()
    
    # SELECIONA O HORARIO
    botao = driver.find_elements_by_class_name('shadow')
    try:
        botao[0].click()
    except:
        print(f'>>> {login} - Nenhuma data disponível.')
        driver.quit()
        return False

    sleep(0.8)
    driver.find_element_by_xpath('//*[@id="confirmacao-horario"]/div[3]/div/button[2]').click()

    # PEGA INFORMAÇÕES DO USUARIO
    id_cliente = driver.execute_script("return sessionStorage.getItem('usuario');")
    id_cliente = loads(id_cliente).get('id')
    token = driver.execute_script("return sessionStorage.getItem('token');")
    header["Authorization"] = f"Bearer {token}"

    if not _TEST:
        # AGUARDA ATÉ 12:00 PARA INICIAR A SOLICITAÇÃO
        print('>>> Aguardando o horário do agendamento')
        comeco = datetime.now().replace(hour=7, minute=59, second=50)
        while datetime.now() < comeco:
            sleep(1)

    s = Session()
    for cookie in driver.get_cookies():
        s.cookies.set(cookie['name'], cookie['value'])

    driver.quit()

    info = {
        "provider": "sefazsp",
        "idServico": "385",
        "data": data_conferida,
    }

    '''info_jales = {
        # JALES
        "idCidade": "63",
        "idUnidade": "52",
        "providedId": "56",
    }'''

    info_campinas = {
        # CAMPINAS
        "idCidade": "24",
        "idUnidade": "13",
        "providedId": "31",
    }

    info_jundiai = {
        # JUNDIAI
        "idCidade": "22",
        "idUnidade": "10",
        "providedId": "160",
    }

    if _CIDADES == 'jundiai':
        info = {**info, **info_jundiai}
    else:
        info = {**info, **info_campinas}

    pagina = s.post(url_busca, json=info, verify=False)
    agenda = pagina.json()
    if not agenda['value']:
        print(f">>> {login} - {data_conferida} indisponivel para agendamento")
        s.close()
        return True

    horarios = agenda['value'][0].get("horarios")
    chave = f"sefazsp{info.get('providedId')}AMBOS"
    finalizado = False
    for hs in horarios:
        info = {
            "chaveConfiguracao": chave,
            "idCliente": id_cliente,
            "data": data_conferida,
            "horario": hs
        }
        while not finalizado:
            try:
                pagina = s.post(url_agenda, headers=header, json=info, verify=False)
                retorno = pagina.json()
                msg = retorno["status"].get("message", '')
                if not msg:
                    print(f'>>> {login} - Agendamento feito para {data_conferida} as {hs}')
                    finalizado = True
                elif msg.startswith("Não é permitido efetuar agendamentos antes de"):
                    if _TEST:
                        finalizado = True
                        print('>>> Rotina executada com sucesso')
                    else:
                        print(">>> Esperando mais um pouco")
                        sleep(2)
                elif msg.startswith("Não foi possível comunicar com o sistema de filas"):
                    print("Não foi possível comunicar com o sistema de filas")
                    break
                elif msg.startswith("Não há vaga disponível para o dia"):
                    print("Não há vaga disponível para o dia")
                    break
                elif msg.startswith("Não foi possível executar a operação"):
                    print(f">>> {login} - Erro de acesso ao servidor.")
                    finalizado = True
                elif msg.startswith("Este cliente já possui 1 agendamento para este serviço nesta mesma data"):
                    print(f">>> {login} - Já possui agendamento para {data_conferida} as {hs}")
                    finalizado = True
                else:
                    print('-'*60)
                    print('>>>>>> Mensagem desconhecida (criar tratamento)')
                    print(login)
                    print(retorno)
                    print('-'*60)
                    break
            except Exception as e:
                print(login, e)
                finalizado = True
        if finalizado:
            break

    s.close()


if __name__ == '__main__':
    dados = {
        "29320051895": "veiga123",  # Robson
        "25194605803": "Veiga123",  # Veiga Jr
        "27329597821": "Veiga123",  # Rodrigo
        "46591958800": "Veiga123",  # Veiga
        "40822327880": "Veiga123",  # Ivonaide
    }

    global _TEST
    global _CIDADES
    if len(sys.argv) <= 1:
        _TEST = input('Rotina de testes (true/false):').lower() == 'true'
        _CIDADES = input('cidades (campinas/jundiai):').lower()
    else:
        _TEST = sys.argv[1].lower() == 'true'
        _CIDADES = "campinas"

    data_validada = data_valida()
    print(">>> Data do agendamento: ", data_validada, _CIDADES)
    if not _TEST:
        inicio = datetime.now().replace(hour=7, minute=58, second=30)
        print(">>> Aguardando até 11:59")
        while datetime.now() < inicio:
            sleep(1)
    else:
        print('>>>>>>>   ROTINA DE TESTE   <<<<<<<')
    
    print(">>> Iniciando Procedimento as ", datetime.now())

    print(">>> Total de threads", len(dados.keys()))
    for key, value in dados.items():
        job = Thread(target=agendar_horario, args=(key, value, data_validada))
        job.start()
        sleep(5)
        if _CIDADES == "jundiai":
            break
