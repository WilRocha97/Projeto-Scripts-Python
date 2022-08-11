# -*- coding: utf-8 -*-
from pandas import read_excel
from time import sleep
from sys import path
import os

path.append(r'..\..\_comum')
from comum_comum import pfx_to_pem, escreve_relatorio_csv
from selenium_comum import exec_phantomjs, send_input, send_select


def get_certificado(certificado):
    _path_cert = r'..\..\_cert'

    if 'AVJ' in certificado:
        return os.path.join(_path_cert, 'CERT AVJ 19142855.pfx'), '19142855'

    elif 'ALDEMAR' in certificado:
        return os.path.join(_path_cert, 'CERT VEIGA 251946.pfx'), '251946'

    elif 'RPEM' in certificado:
        return os.path.join(_path_cert, 'CERT RPEM 35586086.pfx'), '35586086'

    elif 'EVANDRO' in certificado:
        return os.path.join(_path_cert, 'CERT EVANDRO 307828.pfx'), '307828'

    elif 'R. POSTAL' in certificado:
        return os.path.join(_path_cert, 'CERT R POSTAL 26973312.pfx'), '26973312'

    elif 'RODRIGO' in certificado:
        return os.path.join(_path_cert, 'CERT RODRIGO 273295.pfx'), '273295'

    elif 'ROSELEI' in certificado:
        return os.path.join(_path_cert, 'CERT ROSELEI 045838.pfx'), '045838'

    else:
        return os.path.join(_path_cert, 'CERT RPMO 05497487.pfx'), '05497487'


def solicita_extrato(cnpj, certificado, lista_pis):
    url_base = 'https://conectividade.caixa.gov.br'
    url_erro = f'{url_base}/registro/errovoltar.jsp'
    url_troca = f'{url_base}/registro/trocar_raiz.m'

    cert, senha = get_certificado(certificado)
    cnpj = ''.join([n for n in cnpj if n.isdigit()])

    with pfx_to_pem(cert, senha) as cert:
        driver = exec_phantomjs(cert=cert)

    driver.get(url_base)

    # Acessando com o cnpj da empresa
    print('>>> Acessando com o cnpj da empresa')
    driver.get(url_troca)
    sleep(2)

    print('>>> Enviando o input')
    # Pode ser substituido por send_input
    send_input('txt_insc', cnpj, driver)
    driver.execute_script("continuar();")
    sleep(2)

    # aguarda até a página mudar de link
    print('>>> Aguardando a mudança do link')
    if driver.current_url == url_troca:
        count = 0
        while count < 3:
            sleep(2)
            count += 1
            if driver.current_url != url_troca:
                break
        else:
            driver.quit()
            return 'Falha de login com o CNPJ informado'

    print('>>> Verificando o redirecionamento')
    # verifica se foi redirecionado para a página de erro
    if driver.current_url == url_erro:
        driver.quit()
        return 'Certificado invalido para o CNPJ informado'

    # Acessando serviços
    print(">>> Acessando serviços")
    try:
        driver.switch_to_frame(driver.find_element_by_id('frmMenuCNSICP'))
        fnd, slt = ('id', 'cmbAplicativo'), ('text', 'Empregador')
        if not send_select(fnd, slt, driver, delay=1):
            raise Exception()
    except:
        driver.quit()
        return 'Falha ao acessar serviços da empresa'

    # Solicitando Extrato
    try:
        fnd, slt = ('name', 'sltOpcao'), ('text', 'Solicitar Extrato para Fins Recisórios')
        if not send_select(fnd, slt, driver, delay=0.5):
            raise Exception()

        fnd, slt = ('name', 'sltRegiao'), ('value', 'CPD')
        if not send_select(fnd, slt, driver, delay=0.5):
            raise Exception()

        for pis in lista_pis:
            send_input('txtPIS', pis, driver)
            driver.find_element_by_id('adicionar').click()
            sleep(1)

        driver.find_element_by_name('cmdCont').click()
        sleep(1.5)

        print('>>> Solicitação realizada')
    except:
        driver.quit()
        return 'Falha ao solicitar extrato'

    os.makedirs(r'execucao\comprovantes', exist_ok=True)

    caminho = os.path.join(r'execucao\comprovantes', f'{cnpj}.png')
    driver.save_screenshot(caminho)
    driver.quit()

    escreve_relatorio_csv(f"{cnpj};{certificado}", 'solicitadas')
    return 'Executado com sucesso'


def main():
    lista_func = read_excel(r'ignore\funcionarios.xlsx', keep_default_na=False, header=0)
    lista_func.rename(columns={lista_func.columns[0]: 'Cod'}, inplace=True)
    lista_func.query("PIS != ''", inplace=True)
    lista_func = lista_func.groupby(['Cod'])['PIS'].apply(list)

    lista_cuca = read_excel(r'ignore\ExpCli.xls', keep_default_na=False, header=0)
    lista_cuca.rename(columns={lista_cuca.columns[25]: 'ProcRFB'}, inplace=True)
    lista_cuca.query("ProcRFB != ''", inplace=True)

    datas = []
    for cod, lista_pis in lista_func.items():
        emp_cuca = lista_cuca.query(f"Codigo == {cod}")
        try:
            cnpj = emp_cuca.CNPJ.item()
            cert = emp_cuca.ProcRFB.item()
            datas.append((cnpj, cert, lista_pis))
        except:
            pass

    total = len(datas)
    for index, data in enumerate(datas, 1):
        print(f'\n{data[0]} - {index}/{total}')
        result = ''
        try:
            result = solicita_extrato(*data)
        except Exception as e:
            print(e)

        texto = result or 'Ocorreu um erro'
        texto = f"{data[0]};{texto}"
        escreve_relatorio_csv(texto)


if __name__ == '__main__':
    main()
