# -*- coding: utf-8 -*-
import datetime, time, os, pyperclip, re, fitz, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def verificacoes(consulta_tipo, andamento, identificacao, nome):
    alertas = [('inscricao_cancelada.png', 'inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB', False),
               ('nao_foi_possivel.png', 'Não foi possível concluir a consulta', 'erro'),
               ('info_insuficiente.png', 'As informações disponíveis na Secretaria da Receita Federal do Brasil - RFB sobre o contribuinte são insuficientes para a emissão de certidão por meio da Internet.', False),
               ('matriz.png', f'A certidão deve ser emitida para o {consulta_tipo} da matriz', False),
               ('cpf_invalido.png', 'CPF inválido', False),
               ('cnpj_suspenso.png', 'CNPJ suspenso', False),
               ('declaracao_inapta.png', 'CNPJ com situação cadastral declarada inapta pela Secretaria Especial da Receita Federal do Brasil - RFB', False)]
    
    for alerta in alertas:
        if _find_img(alerta[0], conf=0.9):
            if alerta[2] != 'erro':
                _escreve_relatorio_csv(f'{identificacao};{nome};{alerta[1]}', nome=andamento)
            print(f'❌ {alerta[1]}')
            return alerta[2]

    if _find_img('erro_na_consulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)

    return True


def salvar(consulta_tipo, pasta, andamento, identificacao, nome, pasta_download, nome_certidao):
    # espera abrir a tela de salvar o arquivo
    contador = 0
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('servico_indisponivel.png', conf= 0.9):
            return 'erro'
        if _find_img('em_processamento.png', conf=0.9) or _find_img('erro_captcha.png', conf=0.9):
            print('>>> Tentando novamente')
            consulta(consulta_tipo, identificacao)
            if _find_img('site_morreu.png', conf=0.9):
                print('>>> Site morreu, tentando novamente')
                consulta(consulta_tipo, identificacao)
            if contador >= 5:
                _escreve_relatorio_csv('{};{};Consulta em processamento, volte daqui alguns minutos.'.format(identificacao, nome), nome=andamento)
                print(f'❗ Consulta em processamento, volte daqui alguns minutos.')
                p.hotkey('ctrl', 'w')
                return False
            contador += 1
            
        if _find_img('cnpj_vazio.png', conf=0.99):
            p.write(identificacao)
            time.sleep(1)
            
        if _find_img('botao_consultar_selecionado.png', conf=0.9):
            _click_img('botao_consultar_selecionado.png', conf=0.9)
            contador = 0

        if _find_img('nova_certidao.png', conf=0.9):
            _click_img('nova_certidao.png', conf=0.9)

        situacao = verificacoes(consulta_tipo, andamento, identificacao, nome)

        if not situacao:
            p.hotkey('ctrl', 'w')
            return False

        if situacao == 'erro':
            return situacao
        
        if _find_img('cnpj_vazio.png', conf=0.99):
            p.write(identificacao)
            time.sleep(1)
            
        if _find_img(f'informe_{consulta_tipo}.png', conf=0.9):
            p.press('enter')
        
        time.sleep(1)
        
    # escreve o nome do arquivo (.upper() serve para deixar em letra maiúscula)
    time.sleep(1)

    while True:
        try:
            pyperclip.copy(nome_certidao)
            p.hotkey('ctrl', 'v')
            time.sleep(1)
            break
        except:
            print('Erro no clipboard...')
            
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)

    os.makedirs(r'{}\{}'.format(e_dir, pasta), exist_ok=True)
    while True:
        try:
            pyperclip.copy(pasta_download)
            p.hotkey('ctrl', 'v')
            time.sleep(1)
            break
        except:
            print('Erro no clipboard...')

    p.press('enter')
    time.sleep(1)

    p.hotkey('alt', 'l')
    time.sleep(2)

    if _find_img('substituir.png', conf=0.9):
        p.hotkey('alt', 's')

    time.sleep(3)
    return True


def consulta(consulta_tipo, identificacao):
    if consulta_tipo == 'CNPJ':
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir'
    else:
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PF/Emitir'
    
    if not _find_img('receita_logo.png', conf=0.9):
        _abrir_chrome(link)
    
    # espera o site abrir e recarrega caso demore
    while not _find_img(f'informe_{consulta_tipo}.png', conf=0.9):
        if _find_img('nova_certidao.png', conf=0.9):
            _click_img('nova_certidao.png', conf=0.9)
        
        for i in range(60):
            if _find_img('site_morreu.png', conf=0.9):
                break
            time.sleep(1)
            if _find_img(f'informe_{consulta_tipo}.png', conf=0.9):
                break

    time.sleep(1)

    _click_img(f'informe_{consulta_tipo}.png', conf=0.9)
    time.sleep(1)

    p.write(identificacao)
    time.sleep(1)
    p.press('enter')
    return True


def analisa_nome_certidao(pasta_download, nome_certidao):
    arq = os.path.join(pasta_download, nome_certidao)
    doc = fitz.open(arq, filetype="pdf")
    
    texto_arquivo = ''
    for page in doc:
        texto = page.get_text('text', flags=1 + 2 + 8)
        texto_arquivo += texto
    
    nome = re.compile(r'Nome: (.+)').search(texto_arquivo).group(1)
    nome = nome.replace('/', '')
    tem_pendencias = re.compile(r'não constam pendências em seu nome').search(texto_arquivo)
    
    doc.close()
    
    if not tem_pendencias:
        pasta_pendencias = pasta_download + ' com pendências'
        os.makedirs(pasta_pendencias, exist_ok=True)
        shutil.move(arq, os.path.join(pasta_pendencias, nome_certidao.replace('.pdf', f' - {nome}.pdf')))
        return 'com pendência'
    else:
        shutil.move(arq, os.path.join(pasta_download, nome_certidao.replace('.pdf', f' - {nome}.pdf')))
        return ''
    

@_time_execution
@_barra_de_status
def run(window):
    pasta_download = 'V:\Setor Robô\Scripts Python\Receita Federal\Consulta Certidão Negativa\execução\Certidões ' + consulta_tipo
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        identificacao, nome = empresa
        nome = nome.replace('/', '')
        
        nome_certidao = f'Certidão Negativa - {identificacao}.pdf'
        # try:
        while True:
            consulta(consulta_tipo, identificacao)
            situacao = salvar(consulta_tipo, pasta, andamento, identificacao, nome, pasta_download, nome_certidao)
            p.hotkey('ctrl', 'w')

            if not situacao:
                break

            elif situacao == 'erro':
                continue

            else:
                resultado = analisa_nome_certidao(pasta_download, nome_certidao)
                print('✔ Certidão gerada')
                _escreve_relatorio_csv('{};{};{} gerada {}'.format(identificacao, nome, andamento, resultado), nome=andamento)
                time.sleep(1)
                break

        """except:
            time.sleep(2)
            p.hotkey('ctrl', 'w')"""

    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    consulta_tipo = p.confirm(text='Qual o tipo da consulta?', buttons=('CNPJ', 'CPF'))
    pasta = f'Certidões {consulta_tipo}'
    andamento = f'Certidão Negativa {consulta_tipo}'

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
