# -*- coding: utf-8 -*-
from datetime import datetime, date, timedelta
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from bs4 import BeautifulSoup
from requests import Session
import time, os, re, sys
sys.path.append("..")
from comum import pfx_to_pem

'''
Conecta no site 'https://satsp.fazenda.sp.gov.br' utilizando o navegador PhantomJS, realiza o login
com um certificado .pfx e faz o download dos arquivos .xml enviados pelos contribuintes entre as datas
dtInicio e dtFinal, começando em 00:00 e terminando em 23:55
'''

# Verifica se existe alguma mensagem de alerta na tela, retornando False caso encontre
def verifica_alerta(driver):
    try:
        modal = driver.find_element_by_id('dialog-modal')
        print('>>> ', modal.text)
        if modal:
            driver.find_element_by_tag_name('button').click()
            print('>>> Programa Finalizado.')
            return True
    except:
        return False


# Confirma a página acessada, retornando True caso seja diferente do esperado
def verifica_pagina(driver, url):
    if driver.current_url == url:
        print('>>> Página atual: ', driver.current_url)
        return False
    else: 
        print('Erro ao carregar página.', driver.current_url)
        return True


def dados_atualizados(soup):
    viewstate = soup.select("#__VIEWSTATE")[0]['value']
    eventvalidation = soup.select("#__EVENTVALIDATION")[0]['value']

    return (viewstate, eventvalidation)


def pesquisa_lote(dtInicio, equip, driver_cookies, cpfcnpj):
    url_lotes_enviados = "https://satsp.fazenda.sp.gov.br/COMSAT/Private/ConsultarLotesEnviados/PesquisaLotesEnviados.aspx"
    toolkit = ";;AjaxControlToolkit, Version=4.1.50508.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:pt-BR:0c8c847b-b611-49a7-8e75-2196aa6e72fa:de1feab2:fcf0e993:f2c8e708:720a52bf:f9cec9bc:589eaa30:698129cf:fb9b4c57:ccb96cf9"
    horarios = [('00:00', '12:00'), ('12:00', '00:00')]
    regex = re.compile(r'filename=(\d+)\.xml')

    # Inicia uma página para realizar o download dos arquivos
    with Session() as s:
        for cookies in driver_cookies:
            s.cookies.set(cookies['name'], cookies['value'])
        
        for hr in horarios:
            dtFinal = dtInicio if hr != horarios[-1] else dtInicio+timedelta(days=1)
            print(f'>>> Buscando entre {dtInicio} - {hr[0]} e {dtFinal} - {hr[1]}')
            # Acessa a página de consulta dos lotes enviados
            pagina = s.get(url_lotes_enviados, verify=False)
            soup = BeautifulSoup(pagina.content, "html.parser")
            state, validation = dados_atualizados(soup)
           
            info={
                'ToolkitScriptManager1_HiddenField': toolkit, '__EVENTTARGET': '', '__EVENTARGUMENT': '',
                '__VIEWSTATE': state, '__VIEWSTATEGENERATOR': '3008D161', '__SCROLLPOSITIONX': '0',
                '__SCROLLPOSITIONY': '0', '__EVENTVALIDATION': validation, 'ctl00$conteudo$txtNumeroDoRecibo': '',
                'ctl00$conteudo$txtDataInicio': dtInicio.strftime("%d/%m/%Y"), 'ctl00$conteudo$txtHoraInicio': hr[0],
                'ctl00$conteudo$ddlTipoLote': '-2147483648', 'ctl00$conteudo$ddlOrigemLote': '-2147483648',
                'ctl00$conteudo$txtNumeroSerie': str(equip), 'ctl00$conteudo$txtDataTermino': dtFinal.strftime("%d/%m/%Y"),
                'ctl00$conteudo$txtHoraTermino': hr[1], 'ctl00$conteudo$ddlResultadoProcessamento': '-2147483648',
                'ctl00$conteudo$btnPesquisar': 'Pesquisar'
            }

            # Realiza a pesquisa com as informações obtidas anteriormente
            pagina = s.post(url_lotes_enviados, info)

            info.pop('ctl00$conteudo$btnPesquisar')
            proxima = 1
            while(proxima):
                soup = BeautifulSoup(pagina.content, "html.parser")
                # Verifica a quantidade de arquivos disponíveis na página
                arquivos = soup.findAll('a', attrs={'id': re.compile("^conteudo_grvConsultarLotesEnviados_lkbDownloadXml_")})
                if not arquivos:
                    proxima = False
                    continue

                # obtem uma lista do processamento de cada arquivo (sucesso ou falha)
                processados = soup.findAll('span', attrs={'id': re.compile("^conteudo_grvConsultarLotesEnviados_lblResultadoProcessamento_")})

                # Atualiza as informações 
                info['__VIEWSTATE'], info['__EVENTVALIDATION'] = dados_atualizados(soup)
                if info.get('__LASTFOCUS', ''):
                    info.pop('__LASTFOCUS')
                if info.get('ctl00$conteudo$ddlPages', ''):
                    info['ctl00$conteudo$ddlPages'] = str(proxima)

                # Percorre todos os arquivos da página pesquisada
                for x in range(len(arquivos)):
                    if processados[x].text != 'Lote Inválido':
                        item = str(x+2).rjust(2, '0')
                        alvo = f'ctl00$conteudo$grvConsultarLotesEnviados$ctl{item}$lkbDownloadXml'
                        info['__EVENTTARGET'] = alvo

                        salvar = s.post(url_lotes_enviados, info)
                        filename = salvar.headers.get('content-disposition', '')

                        if filename:
                            recibo = regex.search(filename)
                            nome = '_'.join(['sat', cpfcnpj, recibo.group(1), str(equip)])+'.xml'
                            arquivo = open(os.path.join('xml', nome), 'wb')
                            for parte in salvar.iter_content(100000):
                                arquivo.write(parte)
                            arquivo.close()
                            print(f"Arquivo {nome} salvo")

                # Verifica se existe um botão 'Proxima'.
                # Caso exista, a sessão é redirecionada para a próxima página
                try:
                    btnproxima = soup.select("#conteudo_lnkBtnProxima")
                    info['__LASTFOCUS'] = ''
                    info['__EVENTTARGET']= 'ctl00$conteudo$lnkBtnProxima'
                    info['ctl00$conteudo$ddlPages'] = str(proxima)
                    proxima += 1
                    pagina = s.post(url_lotes_enviados, info)
                except:
                    pass

                if not btnproxima: proxima = False


# Realiza o download dos xmls para cada contribuinte vinculado ao certificado .pfx que possui número sat
def consulta_sat(cpfcnpj, certificado, senha, dtInicio, dtFinal):
    url_login = "https://satsp.fazenda.sp.gov.br/COMSAT/Account/LoginSSL.aspx"
    url_selecionar_cnpj = "https://satsp.fazenda.sp.gov.br/COMSAT/Private/SelecionarCNPJ/SelecionarCNPJContribuinte.aspx"
    url_consulta_equip = "https://satsp.fazenda.sp.gov.br/COMSAT/Private/VisualizarEquipamentoSat/VisualizarEquipamentoSAT.aspx"

    cnpj = [cpfcnpj[:8], cpfcnpj[8:]]

    with pfx_to_pem(certificado, senha) as cert:
        driver = webdriver.PhantomJS('phantomjs.exe', service_args=['--ssl-client-certificate-file='+cert])
        driver.set_window_size(1000, 900)
        driver.delete_all_cookies()
        
        # Acessa a página inicial
        try:
            driver.get(url_login)
        except Exception as e:
            print(e)
            driver.quit()
            return False
        time.sleep(1)
        driver.find_element_by_id("conteudo_rbtContabilista").click()
        time.sleep(1)
        driver.find_element_by_id("conteudo_imgCertificado").click()

        # Após logar, acessa a página para selecionar um cnpj
        try:
            driver.get(url_selecionar_cnpj)
        except Exception as e:
            print(e)
            driver.quit()
            return False
        if verifica_pagina(driver, url_selecionar_cnpj): 
            driver.quit()
            return False

        # Inclui os dados do cnpj nos campos de busca e pesquisa
        driver.execute_script(f"document.getElementById('conteudo_txtCNPJ_ContribuinteNro').value = '{cnpj[0]}';")
        driver.execute_script(f"document.getElementById('conteudo_txtCNPJ_ContribuinteFilial').value = '{cnpj[1]}';")
        driver.find_element_by_id('conteudo_btnPesquisar').click()
        time.sleep(2)
        if verifica_alerta(driver): 
            driver.quit()
            return True

        # Seleciona o cnpj restante da pesquisa
        driver.find_element_by_id('conteudo_gridCNPJ_lnkCNPJ_0').click()
        time.sleep(1)
        if verifica_alerta(driver): 
            driver.quit()
            return True

        # Acessa a página de equipamentos SAT
        try:
            driver.get(url_consulta_equip)
        except Exception as e:
            print(e)
            driver.quit()
            return False
        if verifica_pagina(driver, url_consulta_equip): 
            driver.quit()
            return False

        # Seleciona todos os equipamentos ATIVOS (value='10') e pesquisa
        select = Select(driver.find_element_by_id('conteudo_ddlSituacao'))
        try:
            # seleciona pelo "value" do item
            select.select_by_value('10')

            # seleciona pela ordem dos itens
            #select.select_by_index(3)
            # seleciona pelo texto apresentado
            #select.select_by_visible_text('Ativo')
        except Exception as e:
            print('Erro ao selecionar item.')
        time.sleep(2)
        driver.find_element_by_id('conteudo_btnPesquisar').click()
        time.sleep(1)
        if verifica_alerta(driver): 
            driver.quit()
            return True
            
        # Filtrar todos os equipamentos apresentados na pesquisa
        equipamentos = driver.find_elements_by_xpath("//a[starts-with(@id,'conteudo_grvPesquisaEquipamento_lblNumeroSerie')]")
        print('>>> Equipamentos encontrados: ', [t.text for t in equipamentos])
        
        for equipamento in equipamentos:
            print(f'>>> Buscando notas do equipamento {equipamento.text}')
            diferenca = (dtFinal - dtInicio).days + 1
            datas = [dtFinal-timedelta(days=x) for x in range(diferenca)]

            for data in datas:
                pesquisa_lote(data, equipamento.text, driver.get_cookies(), cpfcnpj)
                
        driver.quit()

    return True



if __name__ == '__main__':
    comeco = datetime.now()
    os.makedirs('xml', exist_ok=True)
    #certificado = 'CERT-26973312.pfx'
    #senha = '26973312'
    #certificado = 'CERT-AVJ-19142855.pfx'
    #senha = '19142855'
    certificado = 'CERT-RPEM-307828.pfx'
    senha = '307828'
    #cnpj = "22813534000170" # certificado CERT
    cnpj = "(69092005000198" # certificado 2019
    #cnpj = "29416712000178" # certificado AVJ
    #cnpj = "00196103000172" # sem sat
    #cnpj = "99196103000172" # errado
    dtinicio = date(2020, 3, 3)
    dtfinal = date(2020, 4, 8)

    for x in range(1,4):
        print(f'>>> Tentativa {x}/3')
        try:
            finalizado = consulta_sat(cnpj, certificado, senha, dtinicio, dtfinal)
            if finalizado: break
        except Exception as e:
            print(e)

        print('>>> Ocorreu um erro ao acessar o site. Iniciando nova tentativa.\n')
        time.sleep(2)

    fim = datetime.now()
    print('\n>>> Tempo Total:', (fim-comeco))
