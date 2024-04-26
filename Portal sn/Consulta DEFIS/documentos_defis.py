import datetime, time, os, pyperclip, random, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _get_comp, _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def login(cnpj, cpf, senha):
    def digita_dados_login(cnpj, cpf, senha):
        p.moveTo(p.locateCenterOnScreen(os.path.join('imgs', 'tela_login.png'), confidence=0.9))
        local_mouse = p.position()
        # clica cnpj
        p.click(int(local_mouse[0]), int(local_mouse[1] - 80))
        time.sleep(0.5)
        p.write(cnpj)

        # clica cpf
        p.click(int(local_mouse[0]), int(local_mouse[1] + 20))
        time.sleep(0.5)
        p.write(cpf)

        # clica codigo
        p.click(int(local_mouse[0]), int(local_mouse[1] + 100))
        time.sleep(0.5)
        p.write(senha)
        p.click(50, 280)

    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    while not _find_img('tela_inicial.png', conf=0.95):
        time.sleep(1)

    p.click(50, 280)
    print('>>> Buscando campos para login')
    while not _find_img('tela_login.png', conf=0.8):
        p.click(50, 280)
        p.press('down')

    digita_dados_login(cnpj, cpf, senha)

    p.click(50, 280)
    while not _find_img('continuar.png', conf=0.9):
        if _find_img('codigo_acesso_apagou.png', conf=0.9):
            _click_img('codigo_acesso_apagou.png', conf=0.9)
            time.sleep(0.5)
            p.write(senha)
            time.sleep(1)
            p.click(50, 280)
        p.press('down')

    _click_img('continuar.png', conf=0.9)

    p.click(50, 280)

    while _find_img('carregando.png', conf=0.99):
        time.sleep(1)

    timer = 0
    print('>>> Aguardando tela inicial')
    while not _find_img('bem_vindo.png', conf=0.9):
        if _find_img('captcha_login.png', conf=0.9):
            time.sleep(30)
            if _find_img('bem_vindo.png', conf=0.9):
                break
            return 'recomeçar'
        _click_img('continuar.png', conf=0.9)
        if timer > 10:
            return 'recomeçar'
        if _find_img('tela_inicial.png', conf=0.95):
            _click_position_img('tela_inicial.png', '+', pixels_y=50, conf=0.95)
            print('>>> Verificando ocorrências')
            p.press('end')
            while True:
                _click_img('continuar.png', conf=0.9, timeout=1)
                if _find_img('bem_vindo.png', conf=0.9):
                    break
                if _find_img('codigo_acesso_apagou.png', conf=0.9):
                    _click_img('codigo_acesso_apagou.png', conf=0.9)
                    time.sleep(0.5)
                    p.write(senha)
                    time.sleep(1)
                    p.click(50, 280)
                if _find_img('salvar_senha.png', conf=0.9):
                    p.press('home')
                    break
                p.press('up')
                if _find_img('cod_bloqueado.png', conf=0.9):
                    return 'Código de acesso bloqueado'
                if _find_img('cpf_invalido.png', conf=0.9):
                    return 'CPF não consta como resposável pela pessoa jurídica do CNPJ'
                if _find_img('nao_ha_codigo_para_o_cnpj.png', conf=0.9):
                    return 'Não há código de acesso cadastrado para esse CNPJ'
                if _find_img('codigo_acesso_invalido.png', conf=0.9):
                    return 'Código de acesso inválido'
                if _find_img('campos_obrigatorios.png', conf=0.9):
                    return 'recomeçar'
        time.sleep(1)
        timer += 1

    if _find_img('salvar_senha.png', conf=0.9):
        _click_img('salvar_senha.png', conf=0.9)
        time.sleep(0.5)
        p.press('esc')

    return 'ok'


def busca_defis(cnpj, nome, ano):
    ocorrencia = ''
    print('>>> Buscando a defis')
    # acessa o site do pgdas-defis
    _acessar_site_chrome('https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/defis.app/Captcha.aspx')

    # aguarda a tela do captcha
    while not _find_img('captcha.png', conf=0.9):
        time.sleep(1)

    while True:
        if _find_img('mudar_senha.png', conf=0.9):
            _click_img('mudar_senha.png', conf=0.9)
            time.sleep(0.5)
            p.press('enter')
            ocorrencia = ', Gerenciador de senhas do Google recomenda mudar a senha o quanto antes'
            time.sleep(0.5)
        # clica para resolver o captcha
        _click_img('captcha.png', conf=0.9)
        time.sleep(0.5)

        # aguarda o captcha concluir
        while not _find_img('captcha_resolvido.png', conf=0.9):
            time.sleep(1)
            p.click()

        # aguarda o botão habilitar
        while not _find_img('prosseguir.png', conf=0.9):
            time.sleep(1)

        # clicar em prosseguir
        _click_img('prosseguir.png', conf=0.9)
        time.sleep(0.5)

        # aguarda o login abrir
        if _find_img('selecione_ano_defis.png', conf=0.9):
            break
        if _find_img('continuar_defis.png', conf=0.9):
            break

    print('>>> Selecionando o ano pesquisado')
    # seleciona o ano


    _click_img(f'{ano}.png', conf=0.99)
    time.sleep(1)

    # clica em continuar
    _click_img('continuar_defis.png', conf=0.9)

    # aguarda a tela com o botão para imprimir
    while not _find_img('imprimir.png', conf=0.95):

        if _find_img('nao_tem_defis.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_2.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_3.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_4.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_5.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_6.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_7.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'
        if _find_img('nao_tem_defis_8.png', conf=0.9):
            return f'Ainda não efetuou a DEFIS do ano {ano}'

        time.sleep(1)

    # clica em imprimir
    _click_img('imprimir.png', conf=0.9)

    # aguarda a tela com as defis
    while not _find_img('tabela_defis.png', conf=0.9):
        time.sleep(1)

    # clica em imprimir
    if _find_img(f'defis_{ano}.png', conf=0.99):
        # clica no recibo
        _click_position_img(f'defis_{ano}.png', '+', pixels_x=523, conf=0.99)
        salva_documento('Recibo', cnpj, nome, ano)
        """# clica na declaração
        _click_position_img(f'defis_{ano}.png', '+', pixels_x=598, conf=0.99)
        salva_documento('Declaração', cnpj, nome, ano)"""
        time.sleep(2)
    else:
        return f'Não foi encontrado DEFIS referente a {ano}'

    return f'Documentos DEFIS salvos{ocorrencia}'


def salva_documento(documento, cnpj, nome, ano):
    os.makedirs(f'execução\\{documento}', exist_ok=True)
    print(f'>>> Baixando {documento}')
    caminho = f'V:\Setor Robô\Scripts Python\Portal sn\Salva Declaração Defis\execução\\{documento}'
    arquivo_nome = f'{documento} DEFIS {ano}-{str(int(ano)+ 1)} - {cnpj} - {nome.replace("/", "")}'

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
    print(f'✔ {documento} DEFIS salvo')


@_time_execution
@_barra_de_status
def run(window):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        cnpj, cpf, senha, nome = empresa

        andamentos = f'Documentos Defis'

        # faz login no cpf
        while True:
            _abrir_chrome('https://www8.receita.fazenda.gov.br/SimplesNacional/Servicos/Grupo.aspx?grp=t&area=1',
                          iniciar_com_icone=True)
            resultado = login(cnpj, cpf, senha)
            if resultado != 'recomeçar':
                break
            p.hotkey('ctrl', 'w')

        if resultado == 'ok':
            # busca o informe
            resultado = busca_defis(cnpj, nome, ano)

        _escreve_relatorio_csv(f'{cnpj};{cpf};{nome};{senha};{resultado}', nome=andamentos)
        p.hotkey('ctrl', 'w')
        time.sleep(random.randrange(1, 5))


if __name__ == '__main__':
    ano = _get_comp('0000', '%Y')
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()