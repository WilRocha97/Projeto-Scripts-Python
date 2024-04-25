import time
import os
import pyautogui as p
import PIL.ImagePalette
import fitz
import re
import shutil
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img, _click_position_img, _get_comp
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def seleciona_certificado():
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('botao_gov.png', conf=0.9):
        if _find_img('sessao_finalisada.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            return 'recomeçar'
        time.sleep(1)
    
    # clica no botão do gov.br
    _click_img('botao_gov.png', conf=0.9)
    
    # aguarda o botão de certificado
    while not _find_img('certificado.png', conf=0.95):
        time.sleep(1)
    
    # clica no botão de certificado
    _click_img('certificado.png', conf=0.95)

    timer = 0
    # aguarda a seleção de certificado
    while not _find_img('seleciona_certificado.png'):
        if _find_img('captcha.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return 'recomeçar'
        time.sleep(1)
        timer += 1

        if timer > 30:
            p.hotkey('ctrl', 'w')
            return 'recomeçar'
    
    _click_img('certificado_bianca.png', conf=0.95)
    time.sleep(1)
    
    p.press('enter')


def login(cpf):
    # aguarda a tela inicial
    while not _find_img('selecione_perfil.png', conf=0.9):
        print('>>> Aguardando tela de seleção de CPF')
        time.sleep(1)
    
    if not _find_img('selecione_perfil_ok.png', conf=0.9):
        print('>>> Trocando acesso para CPF')
        _click_position_img('selecione_perfil.png', '+', pixels_y=46, conf=0.9)
        
        while not _find_img('cpf.png', conf=0.9):
            time.sleep(1)
        
        _click_img('cpf.png', conf=0.9)
        
        while not _find_img('informe_cpf.png', conf=0.9):
            print('>>> Aguardando campo para inserir o CPF')
            time.sleep(1)
    
    _click_position_img('informe_cpf.png', '+', pixels_y=46, conf=0.9)
    time.sleep(0.1)
    p.hotkey('ctrl', 'a')
    time.sleep(0.1)
    p.press('delete')
    time.sleep(0.1)

    print('>>> Digitando o CPF')
    p.write(cpf)
    time.sleep(1)
    
    _click_img('verificar.png', conf=0.9)

    timer = 0
    while not _find_img('carregando.png', conf=0.9):
        time.sleep(1)
        timer += 1
        if timer > 5:
            break
    
    while _find_img('carregando.png', conf=0.85):
        time.sleep(1)

    while not _find_img('selecione_o_modulo.png', conf=0.95):
        time.sleep(1)
        if _find_img('erro_campo_acesso.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return False, 'Erro no login, Campo de preenchimento obrigatório'

        if _find_img('erro_procuracao_1.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return False, 'O procurador não possuí perfil com autorização de acesso à Web.'

    return True, 'ok'


def busca_guia(comp, codigo_dominio, cpf, nome, caminho):
    comp_mes, comp_ano = comp.split('/')[0], comp.split('/')[1]
    print('>>> Buscando guia')
    while not _find_img('folha_pagamento.png', conf=0.9):
        _click_img('simplificado.png', conf=0.9, timeout=1)

    while not _find_img('dados_folha_pagamento.png', conf=0.9):
        _click_img('folha_pagamento.png', conf=0.9, timeout=1)

    _click_img('dados_folha_pagamento.png', conf=0.9, timeout=1)

    print('>>> Aguardando tela das guias')
    while not _find_img('tela_folha_pagamento.png', conf=0.9):
        if _find_img('certifique-se.png', conf=0.9):
            _click_img('certifique-se_ok.png', conf=0.95)
        time.sleep(1)

    print('>>> Verificando ano')
    while not _find_img(f'folha_{comp_ano}_marcado.png', conf=0.95):
        if _find_img('certifique-se.png', conf=0.9):
            _click_img('certifique-se_ok.png', conf=0.95)
        while _find_img('carregando.png', conf=0.9):
            time.sleep(1)
        _click_img(f'folha_{comp_ano}.png', conf=0.95, timeout=1)

    print('>>> Verificando mês')
    while not _find_img(f'folha_{comp_mes}_marcado.png', conf=0.98):
        while _find_img('carregando.png', conf=0.9):
            time.sleep(1)
        _click_img(f'folha_{comp_mes}.png', conf=0.98, timeout=1)
    
    if _find_img('sem_valores_recolher.png', conf=0.95):
        return 'Sem trabalhadores e valores a recolher.'
    
    # se não tiver o botão de emitir, fecha a folha.
    if not _find_img('emitir_guia.png', conf=0.95):
        _click_img('encerrar_folha.png', conf=0.95)

        while not _find_img('previa_dae.png', conf=0.9):
            time.sleep(1)

        while not _find_img('confirmar_dae.png', conf=0.9):
            p.press('pgDn')
            time.sleep(1)

        _click_img('confirmar_dae.png', conf=0.95)

        while not _find_img('folha_encerrada.png', conf=0.9):
            time.sleep(1)

        while not _find_img('emitir_guia.png', conf=0.9):
            p.press('pgDn')
            time.sleep(1)

    _click_img('emitir_guia.png', conf=0.9)

    arquivo_nome = f'DAE {comp_mes}-{comp_ano} - Vencimento- - {cpf} - {nome}'
    caminho = os.path.join(caminho, 'Guias')
    resultado_guia = salva_pdf('guia de pagamento', 'DAE gerada com sucesso', arquivo_nome, caminho)

    if not resultado_guia:
        return False

    guia = os.path.join(caminho, arquivo_nome)
    analisa_guia(codigo_dominio, cpf, guia, caminho)

    emitir_recibo_img = 'emitir_recibo.png'
    while not _find_img('emitir_recibo.png', conf=0.9):
        if _find_img('emitir_recibo_2.png', conf=0.9):
            emitir_recibo_img = 'emitir_recibo_2.png'
            break
        p.press('pgDn')
        time.sleep(1)

    _click_img(emitir_recibo_img, conf=0.9)

    arquivo_nome = f'Recibo {comp_mes}-{comp_ano} - {cpf} - {nome}'
    caminho = os.path.join(caminho, 'Recibos')
    resultado_recibo = salva_pdf('recibo', 'Recibo gerado com sucesso', arquivo_nome, caminho)

    return f'{resultado_guia};{resultado_recibo}'


def salva_pdf(documento, documento_andamento, arquivo_nome, caminho):
    os.makedirs(caminho, exist_ok=True)
    print(f'>>> Baixando {documento}')

    # aguarda a tela de salvar do navegador abrir
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_emissao_guia.png', conf=0.9):
            return False
        time.sleep(1)
            
    pyperclip.copy(arquivo_nome)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy(caminho)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('substituir.png', conf=0.95):
        p.press('s')
    time.sleep(1)
    print(f'✔ {documento_andamento}')
    time.sleep(5)

    return documento_andamento


def analisa_guia(codigo_dominio, cpf, guia, caminho):
    texto_arquivo = ''
    guia = guia + '.pdf'
    with fitz.open(os.path.join(caminho, guia)) as pdf:
        for page in pdf:
            texto_pagina = page.get_text('text', flags=1 + 2 + 8)
            texto_arquivo += texto_pagina

    print(texto_arquivo)
    time.sleep(44)

    try:
        cpf_cnpj_guia = re.compile(r'Documento de Arrecadação\ndo eSocial\n(\d\d\d\.\d\d\d\.\d\d\d-\d\d)').search(texto_arquivo).group(1)
        nome_guia = re.compile(r'Documento de Arrecadação\ndo eSocial\n.+\n(.+)').search(texto_arquivo).group(1)
        numero_doc = re.compile(r'Número do Documento\n(.+)').search(texto_arquivo).group(1)
        valor_recolher = re.compile(r'Valor Total do Documento\n(.+,\d+)').search(texto_arquivo).group(1)
        
        vencimento = re.compile(r'Data de Vencimento\n(.+)').search(texto_arquivo).group(1)
        vencimento_sem_barra = vencimento.replace('/', '-')
    except:
        print(texto_arquivo)
        p.alert('ERRO')
        return False, ''
    
    _escreve_relatorio_csv(f'{codigo_dominio};{cpf};{cpf_cnpj_guia};{nome_guia};{numero_doc};{valor_recolher};{vencimento}', nome='Resumo Guias')
    shutil.move(os.path.join(caminho, guia), os.path.join(caminho, guia.replace('Vencimento-', f'Vencimento-{vencimento_sem_barra}')))


@_time_execution
@_barra_de_status
def run(window):
    analisa_guia('', '', 'DAE 04-2024 - Vencimento-20-05-2024 - 02502409802 - LENITA BUCHALLA BAGARELLI', r'\\vpsrv03\Arq_Robo\Esocial\Guias eSocial Domésticas\2024\Ref. 04-2024\Guias')
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = 0
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)
        
        cod_dom, nome, cpf = empresa
        nome = nome.strip()
        while True:
            while True:
                _abrir_chrome('https://login.esocial.gov.br/login.aspx')
                resultados = seleciona_certificado()
                if resultados != 'recomeçar':
                    break

            resultado, mensagem = login(cpf)

            if not resultado:
                _escreve_relatorio_csv(f'{cod_dom};{cpf};{nome};{mensagem}', nome=andamentos)
                break


            resultado = busca_guia(comp, cod_dom, cpf, nome, caminho)
            p.hotkey('ctrl', 'w')
            time.sleep(1)

            if resultado:
                _escreve_relatorio_csv(f'{cod_dom};{cpf};{nome};{resultado}', nome=andamentos)
                break
    
    _escreve_header_csv('CÓD. DOMÍNIO;CPF;NOME;SITUAÇÃO', nome=andamentos)
    _escreve_header_csv('CÓD;CPF;CPF/CNPJ GUIA;NOME;NÚMERO DO DOC.;VALOR;VENCIMENTO', nome='Resumo Guias')
    

if __name__ == '__main__':
    caminho = 'V:\Setor Robô\Scripts Python\Esocial\Gera Guia de Domésticas\execução'
    andamentos = f'Guias eSocial domésticas'
    comp = _get_comp(printable='mm/aaaa', strptime='%m/%Y')

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    
    if index is not None:
        tem_guia = run()
