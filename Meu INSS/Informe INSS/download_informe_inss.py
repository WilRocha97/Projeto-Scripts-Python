import re
import shutil

import pyautogui as p
import time
import os
import pyperclip
import fitz

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def analisa_informe(arquivo):
    texto_arquivo = ''
    with fitz.open(arquivo) as pdf:
        for page in pdf:
            texto_pagina = page.get_text('text', flags=1 + 2 + 8)
            texto_arquivo += texto_pagina

    ano_calendario = re.compile(r'Ano-Calendário (\d{4})').search(texto_arquivo).group(1)
    shutil.move(arquivo, arquivo.replace('Informe de rendimento IR INSS', f'Informe de rendimento IR INSS {ano_calendario}'))


def login(cpf, senha):
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('botao_gov.png', conf=0.95):
        if _find_img('sair.png', conf=0.95):
            _click_img('sair.png', conf=0.95)
        if _find_img('campo_cpf_selecionado.png', conf=0.95):
            break
        
        if _find_img('site_morreu.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'
        
        time.sleep(1)

    # aguarda o campo do cpf
    while not _find_img('campo_cpf_selecionado.png', conf=0.95):
        # clica no botão do gov.br
        _click_img('botao_gov.png', conf=0.95, timeout=1)
        if _find_img('cadastro_incompleto.png', conf=0.95):
            print('>>> Fechando login anterior')
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            _acessar_site_chrome('https://www.gov.br/pt-br')
            print('>>> Aguardando Gov.br')
            while not _find_img('tela_inicial_gov.png', conf=0.95):
                time.sleep(1)

            print('>>> Fechando login')
            if _find_img('entrar_govbr.png', conf=0.95):
                _click_img('entrar_govbr.png', conf=0.95)

            while not _find_img('tela_inicial_gov.png', conf=0.95):
                if _find_img('ja_fechou_usuario.png', conf=0.95):
                    _acessar_site_chrome('https://meu.inss.gov.br/#/login')
                    return 'recomeçar'
                time.sleep(1)

            p.press('tab', presses=14, interval=0.1)
            p.press('enter')
            time.sleep(1)

            while not _find_img('sair_da_conta.png', conf=0.95):
                p.press('pgDn')

            _click_img('sair_da_conta.png', conf=0.95)
            time.sleep(1)
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        if _find_img('alerta_acessos.png', conf=0.9):
            # clica no campo do cpf
            _click_img('alerta_acessos.png', conf=0.9)
            time.sleep(0.5)
            p.press('esc')

        if _find_img('campo_cpf.png', conf=0.9):
            # clica no campo do cpf
            _click_img('campo_cpf.png', conf=0.9)
        
        if _find_img('site_morreu.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'
        
        time.sleep(1)

    time.sleep(0.5)
    print('>>> Digitando o CPF')
    # digita o cpf e aperta enter
    while _find_img('campo_cpf_selecionado.png', conf=0.95):
        p.write(cpf)
        time.sleep(1)
    p.press('enter')

    timer = 0
    # aguarda o campo da senha
    while not _find_img('campo_senha_selecionado.png', conf=0.95):
        if _find_img('erro_403.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        if _find_img('cadastro_divergente_2.png', conf=0.9):
            return 'Foi encontrado uma divergência no cadastro e não será possível realizar o login, caso já tenha sido resolvida junto a RFB, tente novamente mais tarde (ERL0001200)'

        if _find_img('cpf_invalido.png', conf=0.9):
            return 'CPF informado inválido'

        if _find_img('campo_senha.png', conf=0.9):
            # clica no campo da senha
            _click_img('campo_senha.png', conf=0.9)

        """if _find_img('captcha.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'"""
        
        if _find_img('site_morreu.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        time.sleep(1)
        timer += 1

        if timer > 30:
            p.hotkey('ctrl', 'w')
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

    time.sleep(0.5)
    print('>>> Digitando o senha')
    # digita a senha e aperta enter
    while _find_img('campo_senha_selecionado.png', conf=0.95):
        p.write(senha)
        time.sleep(1)
    p.press('enter')

    print('>>> Aguardando a tela inicial')
    # aguarda o site carregar
    while not _find_img('servicos.png', conf=0.9):
        if _find_img('verificacao_duas_etapas_desabilitada.png', conf=0.95):
            _click_img('verificacao_duas_etapas_desabilitada_ok.png', conf=0.95)

        if _find_img('habilitar_dispositivo.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            return 'É necessário utilizar o aplicativo gov.br para autorizar o dispositivo e poder acessar o perfil do usuário'

        if _find_img('cadastro_bloqueado.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            return 'O cadastro do usuário foi bloqueado'

        if _find_img('confirmacao_de_info.png', conf=0.95):
            _click_img('confirmacao_de_info_fechar.png', conf=0.95)
            return 'É necessária confirmação de dados cadastrais'

        if _find_img('erro_403.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        if _find_img('erro_ao_autenticar.png', conf=0.95):
            p.press('enter')
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        if _find_img('autorizacao.png', conf=0.95):
            _click_img('autorizacao_sim.png', conf=0.95)

        if _find_img('termo_gov.png', conf=0.95):
            _click_img('termo_concordar.png', conf=0.95)

        if _find_img('verificacao_duas_etapas.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            return 'É necessária verificação de duas etapas'

        if _find_img('atualizar_cadastro.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            return 'É necessário atualizar o cadastro'

        if _find_img('cadastro_incompleto.png', conf=0.95):
            return 'É necessário atualizar o cadastro'

        if _find_img('usuario_senha_invalido.png', conf=0.95):
            return 'Usuário e/ou senha inválidos'

        if _find_img('captcha.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'
        
        if _find_img('site_morreu.png', conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'

        time.sleep(1)
        if _find_img('botao_gov.png', conf=0.95):
            # clica no botão do gov.br
            _click_img('botao_gov.png', conf=0.95)

    return 'ok'


def busca_informe(ano):
    print('>>> Buscando o Informe')
    # clica em serviços
    _click_img('servicos.png', conf=0.9)

    # aguarda o menu abrir
    while not _find_img('certidoes.png', conf=0.9):
        time.sleep(1)

    # clica em certidões
    _click_img('certidoes.png', conf=0.9)

    # aguarda o menu abrir
    while not _find_img('extrato_ir.png', conf=0.9):
        time.sleep(1)

    # clica em extrato de ir
    _click_img('extrato_ir.png', conf=0.9)

    print('>>> Aguardando o informe')
    # aguarda o extrato aparecer
    while not _find_img('numero_beneficio.png', conf=0.9):
        if _find_img('nao_possui_informe.png', conf=0.9):
            return False, 'Não foi encontrado benefício para consulta'
        if _find_img('cadastro_divergente.png', conf=0.9):
            return False, 'Não foi possível executar o serviço por divergências cadastrais'
        time.sleep(1)

    # configura o ano
    if not _find_img('extrato_ir_disponivel.png', conf=0.9):
        return False, 'Não foi encontrado benefício ativo para gerar informe'

    # clica para abrir o extrato
    informes = p.locateAllOnScreen(os.path.join('imgs', 'extrato_ir_disponivel.png'), confidence=0.9)

    return True, informes


def salva_informe(cpf, nome, contador_informes):
    os.makedirs('execução\Informes IR INSS', exist_ok=True)
    print('>>> Baixando Informe')
    caminho = 'V:\Setor Robô\Scripts Python\Meu INSS\Informe INSS\execução\Informes IR INSS'
    if contador_informes > 1:
        arquivo_nome = f'Informe de rendimento IR INSS - {cpf} - {nome} ({contador_informes})'
    else:
        arquivo_nome = f'Informe de rendimento IR INSS - {cpf} - {nome}'
    # desce a tela até achar o botão de download
    while not _find_img('baixar_pdf.png', conf=0.9):
        p.press('end')

    # clica para baixar o pdf
    _click_img('baixar_pdf.png', conf=0.9)

    # aguarda a tela de salvar do navegador abrir
    _wait_img('salvar_como.png', conf=0.9, timeout=-1)
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
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
    print('✔ Extrato IR INSS baixado com sucesso')
    time.sleep(5)
    analisa_informe(os.path.join(caminho, arquivo_nome + '.pdf'))

    return


def abre_informe(cpf, nome, beneficios):
    contador_informes = 0
    erro_informe = False
    possui_erro = ''
    for informe in beneficios:
        print('>>> Abrindo o informe')
        # aguarda o informe carregar
        timer = 0
        while not _find_img('informe.png', conf=0.9):
            p.click(informe)
            if _find_img('nao_encontrou_informe.png', conf=0.9):
                return 'Não foi encontrado extrato de IR para os dados informados'
            if timer > 5:
                if _find_img('erro_informe.png', conf=0.95):
                    erro_informe = True
                    possui_erro = ', foi encontrado benefício que não gerou informe'
                    break
            time.sleep(1)
            timer += 1

        if not erro_informe:
            # salva o pdf
            contador_informes += 1
            salva_informe(cpf, nome, contador_informes)
            

        # volta para a página dos informes
        p.hotkey('alt', 'left')
        while not _find_img('extrato_ir_disponivel.png', conf=0.9):
            time.sleep(1)

        erro_informe = False

    if contador_informes == 1:
        resultado = f'Extrato IR INSS baixado com sucesso{possui_erro};{contador_informes} Informe encontrado'
    elif contador_informes > 1:
        resultado = f'Extrato IR INSS baixado com sucesso{possui_erro};{contador_informes} Informes encontrados'

    if erro_informe:
        return 'Não foi gerado informe referênte ao benefício encontrado'
    else:
        return resultado


#@_time_execution
#@_barra_de_status
def run():
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index)

        try:
            cpf, nome, senha, ano = empresa
        except:
            ano = ''
            cpf, nome, senha = empresa

        andamentos = f'Informes INSS'

        _abrir_chrome('https://meu.inss.gov.br/#/login')
        # faz login no cpf
        while True:
            resultado = login(cpf, senha)
            if resultado != 'recomeçar':
                break

        if resultado == 'ok':
            # busca o informe
            situacao, resultado = busca_informe(ano)

            if situacao:
                resultado = abre_informe(cpf, nome, resultado)

        _escreve_relatorio_csv(f'{cpf};{nome};{senha};{resultado}', nome=andamentos)
        p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
