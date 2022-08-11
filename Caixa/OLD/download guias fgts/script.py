from Dados import empresas
from selenium import webdriver
from requests.packages import urllib3
from bs4 import BeautifulSoup
from time import sleep
from requests import Session
from datetime import date
import OpenSSL.crypto
import contextlib
import tempfile
import warnings
import os


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings('ignore')

_CERT_ATIVO = ''


def escreve_relatorio_csv(texto):
    try:
        arquivo = open("Resumo.csv", 'a')
    except:
        arquivo = open("Resumo.csv", 'w')
    arquivo.write(texto+'\n')
    arquivo.close()


def atualiza_view(pagina, input_form=False, n=0):
    soup = BeautifulSoup(pagina.content, 'html.parser')
    if input_form:
        viewstate = soup.find('input', attrs={'id': f"j_id1:javax.faces.ViewState:{n}"}).get('value')
    else:
        viewstate = soup.find('update', attrs={'id': f"j_id1:javax.faces.ViewState:{n}"}).text

    return viewstate


def verifica_info(pagina):
    try:
        soup = BeautifulSoup(pagina.content, 'html.parser')
        msg = soup.find('span', attrs={'class': 'ui-messages-info-detail'})
        msg = msg.text.strip()
    except:
        msg = ''

    return msg


def verifica_parcelamentos(pagina):
    obs = ''
    soup = BeautifulSoup(pagina.content, 'html.parser')
    dados = soup.find('update', attrs={'id': 'panelParcelamento'}).text
    soup = BeautifulSoup(dados, 'html.parser')
    try:
        linhas = soup.find('table', attrs={'id': 'tbParcelamento'}).find('tbody')
        for linha in linhas.find_all('tr'):
            cnpj = linha.find('a').text
            if cnpj: obs += cnpj + ' - '
    except Exception as e:
        print(e)

    return obs




def salvar_arquivo(nome, dados):
    os.makedirs('documentos', exist_ok=True)
    arquivo = open(os.path.join('documentos', nome), 'wb')
    for parte in dados.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()
    print(f"Arquivo {nome} salvo")


@contextlib.contextmanager
def pfx_to_pem(pfx_path, pfx_password):
    ''' Decrypts the .pfx file to be used with requests. '''
    with tempfile.NamedTemporaryFile(suffix='.pem', delete=False) as t_pem:
        f_pem = open(t_pem.name, 'wb')
        pfx = open(pfx_path, 'rb').read() # --> type(pfx) > bytes
        p12 = OpenSSL.crypto.load_pkcs12(pfx, pfx_password)
        data_inicial = p12.get_certificate().get_notBefore()
        data_vencimento = p12.get_certificate().get_notAfter()
        if p12.get_certificate().has_expired(): 
            print('Certificado possivelmente vencido')
            data_inicial = p12.get_certificate().get_notBefore()
            data_vencimento = p12.get_certificate().get_notAfter()
            print(f'Inicio: {data_inicial[6:8]}/{data_inicial[4:6]}/{data_inicial[:4]}')
            print(f'Vencimento: {data_vencimento[6:8]}/{data_vencimento[4:6]}/{data_vencimento[:4]}')
        f_pem.write(OpenSSL.crypto.dump_privatekey(OpenSSL.crypto.FILETYPE_PEM, p12.get_privatekey()))
        f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, p12.get_certificate()))
        ca = p12.get_ca_certificates()
        if ca is not None:
            for cert in ca:
                f_pem.write(OpenSSL.crypto.dump_certificate(OpenSSL.crypto.FILETYPE_PEM, cert))
        f_pem.close()
            
        yield t_pem.name


def login(url_login, certificado, senha):
    botao = '//*[@id="content"]/login/div/div[2]/div/div[1]/form[2]/div[2]/div/button'

    with pfx_to_pem(certificado, senha) as cert:
        driver = webdriver.PhantomJS('phantomjs.exe', service_args=['--ssl-client-certificate-file='+cert, '--ignore-ssl-errors=true'])
        driver.delete_all_cookies()
        driver.set_window_size(1000, 900)

        # Acessa a página inicial
        try:
            driver.get(url_login)
            sleep(1)
            driver.find_element_by_xpath(botao).click()
            sleep(1)
        except:
            driver.quit()
            return False

        cookies = driver.get_cookies()
        driver.quit()

    return cookies


def consulta_fgts(vencimento):
    url = 'https://www.conectividadesocial.caixa.gov.br/sifug-web/principal.fug'
    url_consulta = 'https://www.conectividadesocial.caixa.gov.br/sirfg-web/acordo/consultar/consultaAcordoMP927.fug'
    url_detalhes = 'https://www.conectividadesocial.caixa.gov.br/sirfg-web/acordo/consultar/detalheAcordoMP927.fug'
    global _CERT_ATIVO
    hoje = date.today().strftime('%d/%m/%Y')
    mes = vencimento.split('/')[1]
    ano = vencimento.split('/')[2]
    escreve_relatorio_csv('CNPJ;Guia;Inscrições localizadas')

    for index, empresa in enumerate(empresas[:], 1):
        obs = ''
        print(f'>>> Analisando empresa {empresa[0]} - {index}')
        if empresa[1] != _CERT_ATIVO:
            _CERT_ATIVO = empresa[1]
            print('>>> Logando')
            cookies = login(url, empresa[1], empresa[2])
            if not cookies: 
                print('>>> Falha durante login')
                empresas.append(empresa)
                continue
            s = Session()
            for cookie in cookies:
                s.cookies.set(cookie['name'], cookie['value'])

        print('>>> Logou')
        try:
            s.get(url, verify=False)
        except Exception:
            print('>>> Erro de acesso')
            empresas.append(empresa)
            continue
        print('>>> Acessou página')
        sleep(1)
        pagina = s.get(url_consulta, verify=False)

        viewstate = atualiza_view(pagina, True)
        info = {
            'formParcelamento': 'formParcelamento', 'inputConsultaParcelamentoValidator': '',
            'empresa': empresa[0], 'DataTables_Table_0_length': '10',
            'javax.faces.ViewState': viewstate, 'javax.faces.source': 'empresa',
            'javax.faces.partial.event': 'change', 'javax.faces.partial.execute': 'empresa empresa',
            'javax.faces.partial.render': 'formParcelamento',
            'javax.faces.behavior.event': 'change', 'javax.faces.partial.ajax': 'true'
        }
        # Seleciona a empresa para consultar
        pagina = s.post(url_consulta, info, verify=False)
        viewstate = atualiza_view(pagina)
        info = {
            'javax.faces.partial.ajax': 'true', 'javax.faces.source': 'consultarParcelamento',
            'javax.faces.partial.execute': 'consultarParcelamento inputConsultaParcelamentoValidator comboTpInscId inscricaoId',
            'javax.faces.partial.render': 'consultarParcelamento panelParcelamento',
            'consultarParcelamento': 'consultarParcelamento', 'javax.faces.ViewState': viewstate,
            'formParcelamento': 'formParcelamento', 'inputConsultaParcelamentoValidator': '',
            'empresa': empresa[0], 'DataTables_Table_3_length': '10',
        }
        pagina = s.post(url_consulta, info, verify=False)

        obs = verifica_parcelamentos(pagina)
        viewstate = atualiza_view(pagina)
        info = {
            'formParcelamento': 'formParcelamento', 'inputConsultaParcelamentoValidator': '',
            'empresa': empresa[0], 'tbParcelamento_length': '10',
            'DataTables_Table_4_length': '10', 'javax.faces.ViewState': viewstate,
            'j_idt150:0:dadosParcelamento': 'j_idt150:0:dadosParcelamento'
        }
        # Acessa a página das parcelas
        pagina_t = s.post(url_consulta, info, verify=False)
        viewstate_1 = atualiza_view(pagina_t, True)
        info_table = {
            'javax.faces.partial.ajax': 'true', 'javax.faces.source': 'visualizarParcelas',
            'javax.faces.partial.execute': 'visualizarParcelas', 'javax.faces.partial.render': 'panelParcelas',
            'visualizarParcelas': 'visualizarParcelas', 'inputValidatorGeraGuia': '',
            'inputConsultaParcelamentoValidator': '', 'abaSelecionada': '', 'j_idt174': 'j_idt174',
            'j_idt178': hoje, 'javax.faces.ViewState': viewstate_1, 'form': 'form',
            'dataPagamento': '', 'numeroPisMP927': '', 'DataTables_Table_0_length': '10', 'comboMotivo': '',
            'numeroDemanda': '', 'dataProcesso': '', 'numeroProcesso': '', 'dataDeterminacaoAdm': '',
            'areaDemandante': '', 'nomeUnidade': '', 'comboTipoComunicacao': '', 'numeroPisRegularizar': '',
            'comboListaEmpregados': '', 'DataTables_Table_1_length': '10', 'DataTables_Table_2_length': '10',
            'dataRegularizacaoAbaCancelRegularizacao': '', 'competenciaAbaCancelRegularizacao': '',
            'numeroPisAbaCancelRegularizacao': '', 'radioAcaoGerirDados': '1', 'gerirDadosNumeroPisMP927': '',
        }
        # Retorna a tabela com detalhes das parcelas
        pagina = s.post(url_detalhes, info_table, verify=False)
        soup = BeautifulSoup(pagina.content, 'html.parser')
        try:
            content = soup.find('update', attrs={'id': 'panelParcelas'}).text
            soup = BeautifulSoup(content, 'html.parser')
            parcelas = soup.find('table', attrs={'class': 'dataTable'}).find('tbody').find_all('tr')
        except Exception as e:
            print('Erro durante busca de parcelas', e)
            escreve_relatorio_csv(';'.join([empresa[0], 'Erro durante a consulta', obs]))
            parcelas = []

        print('>>> Analisando parcelas')
        for parcela in parcelas:
            colunas = parcela.find_all('td')
            if colunas[2].text.strip() != vencimento: continue
            p = int(colunas[0].text.strip())
            total = len(parcelas)
            # Verificar id do botão da parcela
            item_g = f'j_idt124:0:itemParcela:{p-1}:btnGerarGuiaFgts2'
            item_v = f'j_idt124:0:itemParcela:{p-1}:btnVisualizarGuiaFgts'
            guia_gerada = colunas[3].find('em').get('style')
            viewstate_0 = ''
            if guia_gerada == 'color:red':
                viewstate = atualiza_view(pagina)
                info = {
                    'javax.faces.partial.ajax': 'true', 'javax.faces.source': item_g,
                    'javax.faces.partial.execute': item_g, item_g: item_g, 'form': 'form',
                    'inputValidatorGeraGuia': '', 'DataTables_Table_2_length': '10',
                    'inputConsultaParcelamentoValidator': '', 'javax.faces.ViewState': viewstate,
                    'DataTables_Table_5_length': '10', 'j_idt174': 'j_idt174',
                    'j_idt178': hoje, 'abaSelecionada': '',
                    'javax.faces.ViewState': viewstate_1, 'comboListaEmpregados': '',
                    'numeroPisMP927': '', 'DataTables_Table_0_length': '10', 'comboMotivo': '', 
                    'numeroDemanda': '', 'dataProcesso': '', 'numeroProcesso': '', 
                    'dataDeterminacaoAdm': '', 'areaDemandante': '', 'nomeUnidade': '',
                    'comboTipoComunicacao': '', 'numeroPisRegularizar': '', 'dataPagamento': '',
                    'DataTables_Table_1_length': '10', 'dataRegularizacaoAbaCancelRegularizacao': '',
                    'competenciaAbaCancelRegularizacao': '', 'numeroPisAbaCancelRegularizacao': '',
                    'radioAcaoGerirDados': '1', 'gerirDadosNumeroPisMP927': '',
                }
                # Solicitando a emissão da guia
                pagina = s.post(url_detalhes, info, verify=False)
                viewstate = atualiza_view(pagina)
                # Verificar 'ids' do modal de solicitação de guia para atualizar os campos
                # 'j_idt177' (input), 'j_idt181' (vencimento) e 'j_idt183' (botão)
                info = {
                    'inputValidatorGeraGuia': '', 'dataRegularizacaoAbaCancelRegularizacao': '',
                    'inputConsultaParcelamentoValidator': '', 'abaSelecionada': '',
                    'DataTables_Table_5_length': '10', 'j_idt177': 'j_idt177',
                    'j_idt181': vencimento, 'j_idt183': '', 'javax.faces.ViewState': viewstate_1,
                    'dataPagamento': '', 'numeroPisMP927': '', 'DataTables_Table_0_length': '10',
                    'comboMotivo': '', 'numeroDemanda': '', 'dataProcesso': '', 'form': 'form',
                    'numeroProcesso': '', 'dataDeterminacaoAdm': '', 'areaDemandante': '',
                    'nomeUnidade': '', 'comboTipoComunicacao': '', 'numeroPisRegularizar': '',
                    'comboListaEmpregados': '', 'DataTables_Table_1_length': '10',
                    'competenciaAbaCancelRegularizacao': '', 'radioAcaoGerirDados': '1',
                    'numeroPisAbaCancelRegularizacao': '', 'gerirDadosNumeroPisMP927': '',
                    'DataTables_Table_2_length': '10', 'javax.faces.ViewState': viewstate
                }
                pagina = s.post(url_detalhes, info, verify=False)

                msg = verifica_info(pagina)
                if 'A solicitação de geração de guia foi realizada' not in msg:
                    print(msg)
                    break

                # Retorna para a tela de parcelas
                viewstate_0 = atualiza_view(pagina,True, 1)
                info_table['javax.faces.ViewState'] = viewstate_0
                pagina_t = s.post(url_detalhes, info_table, verify=False)

            if viewstate_0:
                try:
                    viewstate_1 = atualiza_view(pagina_t)
                except:
                    print('>>> Faltou view')
                    viewstate_1 = ''
            else:
                try:
                    viewstate_0 = atualiza_view(pagina_t, True)
                    viewstate_1 = atualiza_view(pagina_t, True, 1)
                except:
                    print('Não encontrou views')
                    viewstate_0 = ''
                    viewstate_1 = ''
            info = {
                'form': 'form', 'inputValidatorGeraGuia': '', 'numeroProcesso': '', 
                'abaSelecionada': '', 'DataTables_Table_5_length': '10', "j_idt174": 'j_idt174',
                'j_idt178': hoje, 'javax.faces.ViewState': viewstate_0, 'dataPagamento': '',
                'numeroPisMP927': '', 'DataTables_Table_0_length': '10', 'comboMotivo': '',
                'numeroDemanda': '', 'dataProcesso': '',  'inputConsultaParcelamentoValidator': '',
                'areaDemandante': '', 'nomeUnidade': '', 'comboTipoComunicacao': '',
                'numeroPisRegularizar': '', 'comboListaEmpregados': '', 'dataDeterminacaoAdm': '',
                'DataTables_Table_1_length': '10', 'dataRegularizacaoAbaCancelRegularizacao': '',
                'competenciaAbaCancelRegularizacao': '', 'numeroPisAbaCancelRegularizacao': '',
                'radioAcaoGerirDados': '1', 'gerirDadosNumeroPisMP927': '', item_v: item_v,
                'DataTables_Table_2_length': '10', 'javax.faces.ViewState': viewstate_1,
            }
            # Download da guia
            guia = s.post(url_detalhes, info, verify=False)
            filename = guia.headers.get("content-disposition", "").split("filename=")

            if filename:
                venc = vencimento.replace('/', '-')
                nome = f'{empresa[0]};IMP_DP;FGTS - Parcelamento;{mes};{ano};{venc};Padrao;Parcela {p}-{total}.pdf'
                salvar_arquivo(nome, guia)
                escreve_relatorio_csv(';'.join([empresa[0], 'Arquivo salvo', obs]))
            else:
                print('>>> Arquivo não localizado')
                escreve_relatorio_csv(';'.join([empresa[0], 'Arquivo não localizado', obs]))
            


if __name__ == '__main__':
    vencimento = '07/12/2020'
    consulta_fgts(vencimento)

