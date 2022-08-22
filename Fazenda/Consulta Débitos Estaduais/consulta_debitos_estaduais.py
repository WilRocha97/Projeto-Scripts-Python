import time

from requests.exceptions import ConnectionError
from bs4 import BeautifulSoup
from sys import path
from requests import Session
from time import sleep

path.append(r'..\..\_comum')
from fazenda_comum import _get_info_post, _new_session_fazenda_driver
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _download_file, _open_lista_dados, _where_to_start, _indice

_situacoes = {
    'C': 'Certidão sem pendencias',
    'S': 'Nao apresentou STDA',
    'G': 'Nao apresentou GIA',
    'E': 'Nao baixou arquivo',
    'T': 'Transporte de Saldo Credor Incorreto',
    'P': 'Pendencias',
    'I': 'Pendencias GIA',
    
}


def confere_pendencias(pagina):
    print('>>> Verificando pendencias')
    
    id_base = 'MainContent_'
    soup = BeautifulSoup(pagina.content, 'html.parser')
    pendencia = [
        # soup.find('span', attrs={'id':f'{id_base}lblMsgErroParcelamento'}).text != '', # parce
        soup.find('span', attrs={'id': f'{id_base}lblMsgErroResultado'}).text != '',  # deb inscritos
        soup.find('span', attrs={'id': f'{id_base}lblMsgErroFiltro'}).text != '',  # ocorrências
    ]
    
    if all(pendencia):
        return _situacoes['C']
    
    situacao = []
    if not pendencia[0]:
        attrs = {'id': f'{id_base}rptListaDebito_lblValorTotalDevido'}
        deb = soup.find('span', attrs=attrs)
        deb = 'R$ 0,00' if not deb else deb.text
        
        pendencia[0] = float(deb[3:].replace('.', '').replace(',', '.')) == 0
        if all(pendencia):
            return _situacoes['C']
        
        situacao.append(_situacoes['P'])
    
    if not pendencia[1]:
        tabela = soup.find('table', attrs={'id': f'{id_base}gdvResultado'})
        if not tabela:
            return 'Erro ao analisar GIA/STDA'
        
        linhas = tabela.find_all('tr')
        if not linhas:
            return 'Erro ao analisar GIA/STDA'
        
        for linha in linhas:
            if 'Não apresentou GIA' in linha.text:
                situacao.append(_situacoes['G'])
                break
            
            if 'Não apresentou STDA' in linha.text:
                situacao.append(_situacoes['S'])
                break
            
            if 'Transporte de Saldo Credor Incorreto' in linha.text:
                situacao.append(_situacoes['T'])
                break
                
            if soup.find('GIA-1/1'):
                situacao.append(_situacoes['I'])
                break
    
    return ' e '.join(situacao)


def consulta_deb_estaduais(cnpj, s, s_id):
    print('>>> Consultando débitos')
    url_base = 'https://www10.fazenda.sp.gov.br/ContaFiscalIcms/Pages'
    url_pesquisa = f'{url_base}/SituacaoContribuinte.aspx?SID={s_id}'
    
    # formata o cnpj colocando os separadores
    f_cnpj = f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
    
    erro = True
    while erro:
        try:
            res = s.get(url_pesquisa)
            time.sleep(2)
            
            state, generator, validation = _get_info_post(content=res.content)
            erro = False
        except:
            erro = True
            
    info = {
        '__EVENTTARGET': 'ctl00$MainContent$btnConsultar',
        '__EVENTARGUMENT': '', '__VIEWSTATE': state,
        '__VIEWSTATEGENERATOR': generator, '__EVENTVALIDATION': validation,
        'ctl00$MainContent$hdfCriterioAtual': '',
        'ctl00$MainContent$ddlContribuinte': 'CNPJ',
        'ctl00$MainContent$txtCriterioConsulta': f_cnpj
    }
    
    erro = True
    while erro:
        try:
            res = s.post(url_pesquisa, info)
            soup = BeautifulSoup(res.content, 'html.parser')
            
            state, generator, validation = _get_info_post(content=res.content)
            erro = False
        except:
            erro = True
    
    info['__EVENTTARGET'] = 'ctl00$MainContent$lkbImpressao'
    info['__VIEWSTATE'] = state
    info['__EVENTVALIDATION'] = validation
    info['__VIEWSTATEGENERATOR'] = generator
    info['ctl00$MainContent$hdfCriterioAtual'] = f_cnpj
    info['ctl00$MainContent$txtCriterioConsulta'] = f_cnpj
    
    id_base = 'MainContent_'
    attrs = {'id': f'{id_base}lblMensagemDeErro'}
    check = soup.find('span', attrs=attrs)
    if check:
        return check.text.strip()
    
    try:
        situacao = confere_pendencias(res)
        if situacao == _situacoes['C']:
            return situacao
    except AttributeError:
        return "Nao identificada"
    
    impressao = s.post(url_pesquisa, info)
    if impressao.headers.get('content-disposition', ''):
        nome = f"{cnpj};INF_FISC_REAL;Debitos Estaduais - {situacao}.pdf"
        _download_file(nome, impressao)
    else:
        situacao = _situacoes['E']
    
    return situacao


@_time_execution
def run():
    # função para abrir a lista de dados
    empresas = _open_lista_dados()
    
    # função para saber onde o robô começa na lista de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    # cria o indice para cada empresa da lista de dados
    total_empresas = empresas[index:]
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    s = False
    
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, usuario, senha, perfil = empresa
        
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa)
        
        # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
        if usuario_anterior != usuario:
            # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
            if s:
                s.close()
            # abre uma nova sessão
            with Session() as s:
                
                erro = 'S'
                contador = 0
                # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
                while erro == 'S':
                    if contador >= 3:
                        cookies = 'erro'
                        sid = 'Erro ao logar na empresa'
                        break
                    try:
                        cookies, sid = _new_session_fazenda_driver(cnpj, usuario, senha, perfil)
                        erro = 'N'
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                        erro = 'S'
                        contador += 1
                
                sleep(1)
                # se não salvar os cookies fecha a sessão e vai para o próximo dado
                if cookies == 'erro':
                    texto = f'{cnpj};{sid}'
                    usuario_anterior = 'padrão'
                    s.close()
                    _escreve_relatorio_csv(texto)
                    print(f'>>> {texto}\n', end='')
                    continue
                
                # adiciona os cookies do login no web driver na sessão por request
                for cookie in cookies:
                    s.cookies.set(cookie['name'], cookie['value'])
        
        # se não retornar a id da sessão do web driver fecha a sessão por request
        if not sid:
            texto = f'{cnpj};Erro no login'
            usuario_anterior = 'padrão'
            s.close()
        
        # se restornar a id da sessão do web driver executa a consulta
        else:
            # retorna o resultado da consulta
            situacao = consulta_deb_estaduais(cnpj, s, sid)
            texto = f'{cnpj};{situacao}'
            # guarda o usuario da execução atual
            usuario_anterior = usuario
        
        # escreve na planilha de andamentos o resultado da execução atual
        _escreve_relatorio_csv(texto)
        print(f'>>> {texto}\n', end='')
    
    # escreve o cabeçalho na planilha de andamentos
    _escreve_header_csv('cnpj;situacao')
    s.close()
    return True


if __name__ == '__main__':
    run()
