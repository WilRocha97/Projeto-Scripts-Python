import datetime, re, shutil, time, os, pyperclip, fitz, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from meu_inss_comum import _login_gov
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


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
        if _find_img('cadastro_inativo.png', conf=0.9):
            return False, 'Cadastro inativo'
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
    elif contador_informes < 1:
        resultado = f'Foi encontrado benefício que não gerou informe'

    return resultado


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


def analisa_informe(arquivo):
    texto_arquivo = ''
    with fitz.open(arquivo) as pdf:
        for page in pdf:
            texto_pagina = page.get_text('text', flags=1 + 2 + 8)
            texto_arquivo += texto_pagina

    ano_calendario = re.compile(r'Ano-Calendário (\d{4})').search(texto_arquivo).group(1)
    shutil.move(arquivo, arquivo.replace('Informe de rendimento IR INSS', f'Informe de rendimento IR INSS {ano_calendario}'))


@_time_execution
@_barra_de_status
def run(window):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        try:
            cpf, nome, senha, ano = empresa
        except:
            ano = ''
            cpf, nome, senha = empresa

        andamentos = f'Informes INSS'

        _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
        # faz login no cpf
        while True:
            resultado = _login_gov(cpf, senha)
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
