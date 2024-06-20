# -*- coding: utf-8 -*-
import datetime, time, os, pyperclip, re, fitz, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def verificacoes(consulta_tipo, andamento, codigo, identificacao, cnpj, nome):
    alertas = [('inscricao_cancelada.png', 'inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB', False),
               ('nao_foi_possivel.png', 'Não foi possível concluir a consulta', 'erro'),
               ('info_insuficiente.png', 'As informações disponíveis na Secretaria da Receita Federal do Brasil - RFB sobre o contribuinte são insuficientes para a emissão de certidão por meio da Internet.', False),
               ('matriz.png', f'A certidão deve ser emitida para o {consulta_tipo} da matriz', False),
               ('cpf_invalido.png', 'CPF inválido', False),
               ('cnpj_invalido.png', 'CNPJ inválido', False),
               ('cnpj_suspenso.png', 'CNPJ suspenso', False),
               ('declaracao_inapta.png', 'CNPJ com situação cadastral declarada inapta pela Secretaria Especial da Receita Federal do Brasil - RFB', False)]
    
    for alerta in alertas:
        if _find_img(alerta[0], conf=0.9):
            if alerta[2] != 'erro':
                if codigo:
                    _escreve_relatorio_csv(f'{codigo};{identificacao};{cnpj};{nome};{alerta[1]}', nome=andamento)
                else:
                    _escreve_relatorio_csv(f'{identificacao};{nome};{alerta[1]}', nome=andamento)
            print(f'❌ {alerta[1]}')
            return alerta[2]

    if _find_img('erro_na_consulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)

    return True


def salvar(consulta_tipo, pasta, andamento, codigo, identificacao, cnpj, nome, pasta_download, nome_certidao):
    # espera abrir a tela de salvar o arquivo
    contador = 0
    timer = 0
    print('>>> Aguardando consulta')
    while not _find_img('salvar_como.png', conf=0.9):
        situacao = verificacoes(consulta_tipo, andamento, codigo, identificacao, cnpj, nome)
        if not situacao:
            p.hotkey('ctrl', 'w')
            return False

        if _find_img('pagina_downloads.png', conf=0.9) or _find_img('site_morreu_2.png', conf=0.9) or _find_img('site_morreu_3.png', conf=0.9) or _find_img('site_morreu_4.png', conf=0.9) or _find_img('site_morreu_5.png', conf=0.9):
            print('>>> Site morreu, tentando novamente')
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            consulta(consulta_tipo, identificacao)

        if situacao == 'erro':
            return situacao

        if _find_img('servico_indisponivel.png', conf= 0.9):
            return 'erro'

        if _find_img('identificacao_nao_informada.png', conf=0.95):
            _click_img('identificacao_nao_informada_ok.png', conf=0.95)
            time.sleep(1)
            p.write(identificacao)
            time.sleep(1)

        if _find_img('em_processamento.png', conf=0.9) or _find_img('em_processamento_2.png', conf=0.9) or _find_img('erro_captcha.png', conf=0.9):
            print('>>> Tentando novamente')
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            consulta(consulta_tipo, identificacao)
            if _find_img('site_morreu_2.png', conf=0.9) or _find_img('site_morreu_3.png', conf=0.9) or _find_img('site_morreu_4.png', conf=0.9):
                print('>>> Site morreu, tentando novamente')
                p.hotkey('ctrl', 'w')
                time.sleep(1)
                consulta(consulta_tipo, identificacao)
            if contador >= 5:
                if codigo:
                    _escreve_relatorio_csv('{};{};{};{};Consulta em processamento, volte daqui alguns minutos.'.format(codigo, identificacao, cnpj, nome), nome=andamento)
                else:
                    _escreve_relatorio_csv('{};{};Consulta em processamento, volte daqui alguns minutos.'.format(identificacao, nome), nome=andamento)

                print(f'❗ Consulta em processamento, volte daqui alguns minutos.')
                p.hotkey('ctrl', 'w')
                return False
            contador += 1

        if _find_img('nova_consulta.png', conf=0.9):
            _click_img('nova_consulta.png', conf=0.9)

        if _find_img('identificacao_nao_informada.png', conf=0.95):
            _click_img('identificacao_nao_informada_ok.png', conf=0.95)
            time.sleep(0.5)
            p.write(identificacao)
            time.sleep(1)
            
        if _find_img(consulta_tipo + '_vazio.png', conf=0.99):
            time.sleep(0.5)
            p.write(identificacao)
            time.sleep(1)
            
        if _find_img('nova_certidao.png', conf=0.9):
            _click_img('nova_certidao.png', conf=0.9)

        if _find_img(f'informe_{consulta_tipo}.png', conf=0.9) or _find_img(f'informe_{consulta_tipo}_2.png', conf=0.9):
            _click_img(f'informe_{consulta_tipo}.png', conf=0.9, timeout=1)
            _click_img(f'informe_{consulta_tipo}_2.png', conf=0.9, timeout=1)
            p.write(identificacao)
            
        if _find_img(consulta_tipo + '_vazio.png', conf=0.99):
            p.write(identificacao)
            time.sleep(1)
            p.press('enter')
            
        if _find_img('botao_consultar_selecionado.png', conf=0.9):
            _click_img('botao_consultar_selecionado.png', conf=0.9)
            contador = 0

        time.sleep(1)
        timer += 1

        if timer > 30:
            print('>>> Site morreu, tentando novamente')
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            consulta(consulta_tipo, identificacao)
            timer = 0

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
        p.click(900, 750)
        p.click(934, 750)
        if _find_img('nova_certidao.png', conf=0.9):
            _click_img('nova_certidao.png', conf=0.9)

        if _find_img('nova_consulta.png', conf=0.9):
            _click_img('nova_consulta.png', conf=0.9)

        if _find_img('site_bugou.png', conf=0.9):
            p.press('f5')

        for i in range(60):
            if _find_img('site_morreu_2.png', conf=0.9) or _find_img('site_morreu_3.png', conf=0.9) or _find_img('site_morreu_4.png', conf=0.9) or _find_img('site_morreu_6.png', conf=0.9) or _find_img('pagina_downloads.png', conf=0.9):
                p.hotkey('ctrl', 'w')
                _abrir_chrome(link)

            time.sleep(1)
            if _find_img(f'informe_{consulta_tipo}.png', conf=0.9) or _find_img(f'informe_{consulta_tipo}_2.png', conf=0.9):
                break

    time.sleep(1)

    _click_img(f'informe_{consulta_tipo}.png', conf=0.9, timeout=1)
    _click_img(f'informe_{consulta_tipo}_2.png', conf=0.9, timeout=1)
    time.sleep(1)

    p.write(identificacao)
    time.sleep(1)
    p.press('enter')
    return True


def analisa_nome_certidao(pasta_download, nome_certidao):
    while True:
        try:
            if _find_img('receita_logo.png', conf=0.9):
                p.hotkey('ctrl', 'w')

            arq = os.path.join(pasta_download, nome_certidao)
            try:
                doc = fitz.open(arq, filetype="pdf")
            except:
                return 'erro'

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

        except PermissionError:
            print('❗ Erro ao analizar o arquivo, tentando novamente...')
            _click_img('continuar_download.png', conf=0.9)


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
        
        try:
            codigo, identificacao, cnpj, nome = empresa
        except:
            try:
                codigo, identificacao, nome = empresa
                cnpj = False
            except:
                identificacao, nome = empresa
                codigo = False
                cnpj = False
            
        nome = nome.replace('/', '')
        
        if codigo:
            nome_certidao = f'Certidão Negativa - {codigo} - {identificacao}.pdf'
        else:
            nome_certidao = f'Certidão Negativa - {identificacao}.pdf'
            
        # try:
        while True:
            consulta(consulta_tipo, identificacao)
            situacao = salvar(consulta_tipo, pasta, andamento, codigo, identificacao, cnpj, nome, pasta_download, nome_certidao)
            p.hotkey('ctrl', 'w')

            if not situacao:
                break

            elif situacao == 'erro':
                continue

            else:
                resultado = analisa_nome_certidao(pasta_download, nome_certidao)
                if resultado == 'erro':
                    p.hotkey('ctrl', 'w')
                else:
                    print('✔ Certidão gerada')
                    if codigo:
                        _escreve_relatorio_csv('{};{};{};{};{} gerada {}'.format(codigo, identificacao, cnpj, nome, andamento, resultado), nome=andamento)
                    else:
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
