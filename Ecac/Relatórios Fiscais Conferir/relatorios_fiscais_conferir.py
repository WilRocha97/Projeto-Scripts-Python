# -*- coding: utf-8 -*-
import re, fitz, shutil, time, pyperclip, os, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _click_img, _wait_img, _find_img, _click_position_img, _click_position_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Confere IR.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


def abrir_chrome():
    p.hotkey('win', 'm')

    if _find_img('chrome_aberto.png', conf=0.99):
        _click_img('chrome_aberto.png', conf=0.99, timeout=1)
    else:
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        while not _find_img('google.png', conf=0.9):
            time.sleep(1)
            p.moveTo(1163, 377)
            p.click()

    link = 'https://portal.conferironline.com.br/dashboard'

    _click_img('maxi.png', conf=0.9, timeout=1)
    p.click(1000, 51)
    time.sleep(1)
    p.write(link.lower())
    time.sleep(1)
    p.press('enter')
    print('>>> Abrindo o site do Conferir')
    return 'ok'


def login_conferir():
    print('>>> Logando no Conferir')
    # email
    _click_img('campo_email.png', conf=0.9)
    time.sleep(0.5)
    p.hotkey('ctrl', 'a')
    p.write(user[0])
    
    # senha
    _click_img('campo_senha.png', conf=0.9)
    time.sleep(0.5)
    p.hotkey('ctrl', 'a')
    p.write(user[1])
    
    time.sleep(0.5)
    p.press('enter')
    

def consulta(cpf):
    print('>>> Aguardando o site Conferir')
    # aguarda barra de pesquisa
    while not _find_img('barra_de_pesquisa.png', conf=0.9):
        if _find_img('tela_login.png', conf=0.9):
            login_conferir()
        time.sleep(1)
    
    print('>>> Buscando CPF')
    # clica na barra de pesquisa
    _click_img('barra_de_pesquisa.png', conf=0.9)
    time.sleep(1)
    
    # escreve o cpf
    p.write(cpf)
    
    # aguarda o CPF consultado aparecer
    while not _find_img('empresa.png', conf=0.9):
        if _find_img('nenhum_resultado_encontrado.png', conf=0.9):
            print('❗ CPF não encontrado no sistema')
            return 'CPF não encontrado no sistema'
        time.sleep(1)
    
    # abre o perfil do CPF
    _click_img('empresa.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o menu do ecac aparecer
    while not _find_img('acoes_ecac.png', conf=0.9):
        time.sleep(1)
    
    # clica no menu do ecac
    _click_img('acoes_ecac.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o botão de CND no ecac aparecer, se aparecer uma mensagem sobre login inválido retorna o erro
    while not _find_img('cnd_ecac.png', conf=0.9):
        if _find_img('login_invalido.png', conf=0.9):
            print('❗ Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.')
            return 'Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.'
        time.sleep(1)
    
    print('>>> Acessando ECAC')
    # clica no botão de CND do ecac
    _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)

    resultado = gera_relatorio()
    if resultado != 'ok':
        return resultado
    
    # aguarda a tela para salvar o PDF
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_sistema_5.png', conf=0.9) or _find_img('erro_sistema_8.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            time.sleep(1)
            p.press('f5')
            time.sleep(2)

            while not _find_img('cnd_ecac.png', conf=0.9):
                time.sleep(1)

            _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)

            resultado = gera_relatorio()
            if resultado != 'ok':
                return resultado

        time.sleep(1)

    time.sleep(2)
    return 'ok'


def gera_relatorio():
    # aguarda o botão de emissão da certidão aparecer
    timer = 0
    print('>>> Aguardando ECAC')
    while not _find_img('gerar_relatorio.png', conf=0.9):
        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('erro_sistema_3.png', conf=0.9):
            print('❗ Solicitação rejeitada pelo sistema, tente novamente mais tarde.')
            return 'Solicitação rejeitada pelo sistema, tente novamente mais tarde.'
            
        # se aparecer a tela de consulta em processamento, retorna o erro
        if _find_img('erro_sistema_7.png', conf=0.9):
            print('❗ Consulta não liberada pelo sistema, tente mais tarde.')
            return 'Consulta não liberada pelo sistema, tente mais tarde.'
            
        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('mensagens_importantes_ecac.png', conf=0.9):
            print(
                '❗ Este CPF possuí mensagens importantes no ECAC, não é possível emitir o relatório até que a mensagem seja aberta.')
            return 'Este CPF possuí mensagens importantes no ECAC, não é possível emitir o relatório até que a mensagem seja aberta.'

        # se demorar 5 segundos para o botão de emissão da certidão aparecer, verifica se a tela de login do ecac stá bugada
        # se a tela de login do ecac estiver bugada, fecha a janela, recarrega a página no conferir e clica no botão de CND do ecac novamente
        if timer > 10:
            if _find_img('erro_sistema.png', conf=0.9)\
                    or _find_img('erro_sistema_2.png', conf=0.9)\
                    or _find_img('erro_sistema_4.png', conf=0.9)\
                    or _find_img('erro_sistema_6.png', conf=0.9)\
                    or _find_img('erro_sistema_8.png', conf=0.9):
                print('>>> Erro no ECAC, tentando novamente')
                p.hotkey('ctrl', 'w')
                time.sleep(1)
                p.press('f5')
                time.sleep(2)

                while not _find_img('cnd_ecac.png', conf=0.9):
                    time.sleep(1)

                _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)
                timer = 0

        if timer > 60:
            # se demorar muito pata carregar, retorna o erro
            if _find_img('sistema_carregando.png', conf=0.9):
                print('❌ Erro ao acessar o ECAC, sistema demorou muito para responder.')
                return 'Erro ao acessar o ECAC, sistema demorou muito para responder.'

        time.sleep(1)
        timer += 1

    time.sleep(2)
    # clica para abrir a tela do relatório
    print('>>> Gerando relatório')
    _click_img('gerar_relatorio.png', conf=0.9)
    
    timer = 0
    while not _find_img('botao_gerar_relatorio.png', conf=0.9):
        if timer > 10:
            _click_img('gerar_relatorio_clicado.png', conf=0.9)
            # se aparecer essa mensagem de erro, retorna 'ok' para o erro seja tratado dentro do while anterior
            if _find_img('erro_sistema_8.png', conf=0.9):
                print('>>> Erro no ECAC, tentando novamente')
                return 'ok'
        time.sleep(1)
        timer += 1
    
    # clica no botão para emitir o relatório
    _click_img('botao_gerar_relatorio.png', conf=0.9)

    return 'ok'


def salvar_pdf(pasta_final):
    print('>>> Salvando relatório')
    while not _find_img('salvar_como.png', conf=0.9):
        time.sleep(1)
        
    os.makedirs(pasta_final, exist_ok=True)
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)

    erro = 'sim'
    while erro == 'sim':
        try:
            pyperclip.copy(pasta_final)
            p.hotkey('ctrl', 'v')
            erro = 'não'
        except:
            erro = 'sim'

    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    while _find_img('salvar_como.png', conf=0.9):
        if _find_img('substituir.png', conf=0.9):
            p.press('s')
    time.sleep(1)
    p.hotkey('ctrl', 'w')

    print('✔ Relatório emitido com sucesso')
    return 'Relatório emitido com sucesso'


def verifica_relatorio(pasta_analise, pasta_final, pasta_final_sem_pendencias):
    print('>>> Analisando relatório')
    # Analisa cada pdf que estiver na pasta
    for arquivo in os.listdir(pasta_analise):
        print(f'Arquivo: {arquivo}')
        # Abrir o pdf
        arq = os.path.join(pasta_analise, arquivo)
    
        with fitz.open(arq) as pdf:
            # Para cada página do pdf, se for a segunda página o script ignora
            for count, page in enumerate(pdf):
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                if re.compile(r'Diagnóstico Fiscal na Receita Federal e Procuradoria-Geral da Fazenda Nacional.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                    os.makedirs(pasta_final_sem_pendencias, exist_ok=True)
                    situacao = 'Sem pendências'
                    pasta_final = os.path.join(pasta_final_sem_pendencias, arquivo.replace('.pdf', '-Sem pendências.pdf'))
                    break
                else:
                    os.makedirs(pasta_final, exist_ok=True)
                    if re.compile(r'Diagnóstico Fiscal na Receita Federal').search(textinho):
                        situacao_2 = ''
                        if not re.compile(r'Diagnóstico Fiscal na Receita Federal.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                            situacao_1 = 'Pendência na Receita Federal'
                        else:
                            situacao_1 = 'Sem pendência na Receita Federal'
                            
                    if re.compile(r'Diagnóstico Fiscal na Procuradoria-Geral da Fazenda Nacional').search(textinho):
                        situacao_1 = ''
                        if not re.compile(r'Diagnóstico Fiscal na Procuradoria-Geral da Fazenda Nacional.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                            situacao_2 = 'Pendência na Procuradoria Geral da Fazenda Nacional'
                        else:
                            situacao_2 = 'Sem pendencias na Procuradoria Geral da Fazenda Nacional'
                    
                    situacao = f'Com pendências;{situacao_1};{situacao_2}'
                    
        shutil.move(arq, pasta_final)
    
    print(situacao.replace(';', ' | '))
    return situacao


@_time_execution
def run():
    pasta_analise = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\ignore\Relatórios'
    pasta_final = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\execução\Relatórios'
    pasta_final_sem_pendencias = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\execução\Relatórios Sem Pendências'
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # configurar o índice para localizar em qual empresa está
        _indice(count, total_empresas, empresa, index)
        cpf, nome = empresa
        
        abrir_chrome()
        resultado = consulta(cpf)
        if resultado == 'ok':
            resultado = salvar_pdf(pasta_analise)
            situacao = verifica_relatorio(pasta_analise, pasta_final, pasta_final_sem_pendencias)
        else:
            situacao = ''

        p.hotkey('ctrl', 'w')
        _escreve_relatorio_csv(f'{cpf};{nome};{resultado};{situacao}')


if __name__ == '__main__':
    run()
