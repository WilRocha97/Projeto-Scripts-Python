import datetime, random, re, shutil, time, os, pyperclip, fitz, pyautogui as p
from bs4 import BeautifulSoup

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img

imagens = 'V:\\_imagens_comum_python\\imgs_comum_meu_inss'


def login_gov(cpf, senha):
    # aguarda o site carregar
    print('>>> Aguardando o site carregar')
    p.moveTo(5, 5)
    time.sleep(random.randrange(1, 5))
    while not _find_img('botao_gov.png', pasta=imagens, conf=0.95):
        if _find_img('fechar_acesso.png', pasta=imagens, conf=0.95):
            _click_img('fechar_acesso.png', pasta=imagens, conf=0.95)

        if _find_img('sair_fgts.png', pasta=imagens, conf=0.95):
            _click_img('sair_fgts.png', pasta=imagens, conf=0.95)
        
        if _find_img('sair.png', pasta=imagens, conf=0.95):
            _click_img('sair.png', pasta=imagens, conf=0.95)
        
        if _find_img('campo_cpf_selecionado.png', pasta=imagens, conf=0.95):
            break
        
        if _find_img('site_morreu.png', pasta=imagens, conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
            return 'recomeçar'
        
        time.sleep(1)
    
    # aguarda o campo do cpf
    time.sleep(random.randrange(1, 5))
    while not _find_img('campo_cpf_selecionado.png', pasta=imagens, conf=0.95):
        if _find_img('sair.png', pasta=imagens, conf=0.95):
            _click_img('sair.png', pasta=imagens, conf=0.95)
        # clica no botão do gov.br
        p.moveTo(5, 5)
        _click_img('botao_gov.png', pasta=imagens, conf=0.95, timeout=1)
        if _find_img('cadastro_incompleto.png', pasta=imagens, conf=0.95):
            print('>>> Fechando login anterior')
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            _acessar_site_chrome('https://www.gov.br/pt-br')
            print('>>> Aguardando Gov.br')
            while not _find_img('tela_inicial_gov.png', pasta=imagens, conf=0.95):
                time.sleep(1)
            
            print('>>> Fechando login')
            if _find_img('entrar_govbr.png', pasta=imagens, conf=0.95):
                _click_img('entrar_govbr.png', pasta=imagens, conf=0.95)
            
            while not _find_img('tela_inicial_gov.png', pasta=imagens, conf=0.95):
                if _find_img('ja_fechou_usuario.png', pasta=imagens, conf=0.95):
                    _acessar_site_chrome('https://meu.inss.gov.br/#/login')
                    return 'recomeçar'
                time.sleep(1)
            
            p.press('tab', presses=14, interval=0.1)
            p.press('enter')
            time.sleep(1)
            
            while not _find_img('sair_da_conta.png', pasta=imagens, conf=0.95):
                p.press('pgDn')
            
            _click_img('sair_da_conta.png', pasta=imagens, conf=0.95)
            time.sleep(1)
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'
        
        if _find_img('alerta_acessos.png', pasta=imagens, conf=0.9):
            # clica no campo do cpf
            _click_img('alerta_acessos.png', pasta=imagens, conf=0.9)
            time.sleep(0.5)
            p.press('esc')
        
        if _find_img('campo_cpf.png', pasta=imagens, conf=0.9):
            # clica no campo do cpf
            _click_img('campo_cpf.png', pasta=imagens, conf=0.9)
        
        if _find_img('site_morreu.png', pasta=imagens, conf=0.95):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.hotkey('ctrl', 'w')
            _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
            return 'recomeçar'
        
        time.sleep(1)
    
    time.sleep(0.5)
    print('>>> Digitando o CPF')
    # digita o cpf e aperta enter
    while _find_img('campo_cpf_selecionado.png', pasta=imagens, conf=0.95):
        p.write(cpf)
        time.sleep(1)
    p.press('enter')
    
    timer = 0
    # aguarda o campo da senha
    while not _find_img('campo_senha_selecionado.png', pasta=imagens, conf=0.95):
        erros = [('campo_senha.png', False, 'campo_senha.png', False, False, 0, False, False),
                 ('cadastro_divergente_2.png', 'Foi encontrado uma divergência no cadastro e não será possível realizar o login, caso já tenha sido resolvida junto a RFB, tente novamente mais tarde (ERL0001200)', False, False, False, 0, False, False),
                 ('cpf_invalido.png', 'CPF informado inválido', False, False, False, 0, False, False),
                 ('confirmacao_de_info_2.png', 'É necessária confirmação de dados cadastrais', False, False, False, 0, False, False),
                 ('erro_403.png', 'recomeçar', False, False, True, 1, True, False),
                 ('sistema_indisponivel.png', 'recomeçar', False, False, True, 1, True, False),
                 ('site_morreu.png', 'recomeçar', False, False, True, 2, False, True),
                 ('erro_gov.png', 'recomeçar', False, False, True, 2, False, True), ]
        for erro in erros:
            if _find_img(erro[0], pasta=imagens, conf=0.95):
                p.moveTo(5, 5)
                if erro[2]:
                    _click_img(erro[2], pasta=imagens, conf=0.95)
                if erro[3]:
                    p.press('enter')
                if erro[4]:
                    for i in range(int(erro[5])):
                        p.hotkey('ctrl', 'w')
                        time.sleep(1)
                if erro[6]:
                    _acessar_site_chrome('https://meu.inss.gov.br/#/login')
                if erro[7]:
                    _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
                if erro[1]:
                    return erro[1]
        
        time.sleep(random.randrange(1, 5))
        timer += 1
        if timer > 30:
            p.hotkey('ctrl', 'w')
            _acessar_site_chrome('https://meu.inss.gov.br/#/login')
            return 'recomeçar'
        
        timer_2 = 0
        while _find_img('captcha.png', pasta=imagens, conf=0.95):
            time.sleep(1)
            timer_2 += 1
            
            if timer_2 > 60:
                p.hotkey('ctrl', 'w')
                time.sleep(1)
                p.hotkey('ctrl', 'w')
                _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
                return 'recomeçar'
    
    time.sleep(0.5)
    print('>>> Digitando o senha')
    # digita a senha e aperta enter
    while _find_img('campo_senha_selecionado.png', pasta=imagens, conf=0.95):
        p.write(senha)
        time.sleep(1)
    p.press('enter')
    
    print('>>> Aguardando a tela inicial')
    # aguarda o site carregar
    while not _find_img('servicos.png', pasta=imagens, conf=0.9):
        erros_2 = [('verificacao_duas_etapas_desabilitada.png', False, 'verificacao_duas_etapas_desabilitada_ok.png', False, False, 0, False, False),
                   ('habilitar_dispositivo.png', 'É necessário utilizar o aplicativo gov.br para autorizar o dispositivo e poder acessar o perfil do usuário', False, False, True, 1, False, False),
                   ('cadastro_bloqueado.png', 'O cadastro do usuário foi bloqueado', False, False, True, 1, False, False),
                   ('confirmacao_de_info_3.png', 'É necessária confirmação de dados cadastrais', False, False, False, 0, False, False),
                   ('confirmacao_de_info.png', 'É necessária confirmação de dados cadastrais', 'confirmacao_de_info_fechar.png', False, False, 0, False, False),
                   ('erro_403.png', 'recomeçar', False, False, True, 1, True, False),
                   ('erro_ao_autenticar.png', 'recomeçar', False, True, False, 0, True, False),
                   ('erro_ao_autenticar_2.png', 'recomeçar', False, True, False, 0, True, False),
                   ('pagina_sem_resposta.png', 'recomeçar', False, True, True, 1, True, False),
                   ('autorizacao.png', False, 'autorizacao_sim.png', False, False, 0, False, False),
                   ('termo_gov.png', False, 'termo_concordar.png', False, False, 0, False, False),
                   ('verificacao_duas_etapas.png', 'É necessária verificação de duas etapas', False, False, True, 1, False, False),
                   ('atualizar_cadastro.png', 'É necessário atualizar o cadastro', False, False, True, 1, False, False),
                   ('cadastro_incompleto.png', 'É necessário atualizar o cadastro', False, False, False, 0, False, False),
                   ('usuario_senha_invalido.png', 'Usuário e/ou senha inválidos', False, False, False, 0, False, False),
                   ('cpf_nao_encontrado.png', 'O CPF informado não foi localizado na base de dados do INSS (CNIS)', False, False, False, 0, False, False),
                   ('site_morreu.png', 'recomeçar', False, False, True, 2, True, True),
                   ('botao_gov.png', False, 'botao_gov.png', False, False, 0, False, False)]
        for erro_2 in erros_2:
            if _find_img(erro_2[0], pasta=imagens, conf=0.95):
                p.moveTo(5, 5)
                if erro_2[2]:
                    _click_img(erro_2[2], pasta=imagens, conf=0.95)
                
                if erro_2[3]:
                    p.press('enter')
                
                if erro_2[4]:
                    for i in range(int(erro_2[5])):
                        p.hotkey('ctrl', 'w')
                        time.sleep(1)
                
                if erro_2[6]:
                    _acessar_site_chrome('https://meu.inss.gov.br/#/login')
                
                if erro_2[7]:
                    _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
                
                if erro_2[1]:
                    return erro_2[1]
        
        timer = 0
        while _find_img('captcha.png', pasta=imagens, conf=0.95):
            time.sleep(1)
            timer += 1
            
            if timer > 60:
                p.hotkey('ctrl', 'w')
                time.sleep(1)
                p.hotkey('ctrl', 'w')
                _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
                return 'recomeçar'
        
        time.sleep(random.randrange(1, 5))
    return 'ok'
_login_gov = login_gov