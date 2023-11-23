# -*- coding: utf-8 -*-
import re, fitz, shutil, time, pyperclip, os, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _click_img, _wait_img, _find_img, _click_position_img, _click_position_img
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Confere IR.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


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
    time.sleep(3)
    

def busca_cpf(cpf):
    print('>>> Aguardando o site Conferir')
    # aguarda barra de pesquisa
    timer = 0
    while not _find_img('barra_de_pesquisa.png', conf=0.9):
        time.sleep(1)
        if _find_img('tela_login.png', conf=0.9):
            login_conferir()
            time.sleep(3)
            
        if timer > 60:
            p.press('f5')
            timer = 0
            
        timer += 1
    
    print('>>> Buscando CPF')
    # clica na barra de pesquisa
    _click_img('barra_de_pesquisa.png', conf=0.9)
    time.sleep(1)
    
    # escreve o cpf
    p.write(cpf)
    
    # aguarda o CPF consultado aparecer
    while not _find_img('empresa.png', conf=0.9):
        if _find_img('nenhum_resultado_encontrado.png', conf=0.9):
            return '❗ CPF não encontrado no sistema'
        time.sleep(1)
    
    # abre o perfil do CPF
    _click_img('empresa.png', conf=0.9)
    time.sleep(1)
    return '✔ OK'


def consulta_relatorio():
    # aguarda o menu do ecac aparecer
    while not _find_img('acoes_ecac.png', conf=0.9):
        if _find_img('cookies.png', conf=0.9):
            _click_img('entendi_cookies.png', conf=0.9)
        time.sleep(1)
    
    # clica no menu do ecac
    _click_img('acoes_ecac.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o botão de CND no ecac aparecer, se aparecer uma mensagem sobre login inválido retorna o erro
    while not _find_img('cnd_ecac.png', conf=0.9):
        if _find_img('login_invalido.png', conf=0.9):
            return '❗ Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.'
        if _find_img('cnd_ecac_2.png', conf=0.9):
            break
        p.press('PgDn')
        time.sleep(1)
    
    print('>>> Acessando ECAC')
    # clica no botão de CND do ecac
    _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)
    _click_position_img('cnd_ecac_2.png', '+', pixels_y=50, conf=0.9)
    
    resultado = gera_relatorio()
    if resultado != 'ok':
        return resultado

    time.sleep(2)
    return '✔ OK'


def reiniciar_ecac():
    print('>>> Erro no ECAC, tentando novamente')
    p.click(1000, 10)
    time.sleep(1)
    p.hotkey('ctrl', 'w')
    time.sleep(1)
    p.press('f5')
    time.sleep(2)
    
    while not _find_img('cnd_ecac.png', conf=0.9):
        time.sleep(1)
    
    while _find_img('cnd_ecac.png', conf=0.9):
        _click_position_img('cnd_ecac.png', '+', pixels_y=92, conf=0.9)
        time.sleep(1)
    


def gera_relatorio():
    # aguarda o botão de emissão da certidão aparecer
    timer = 0
    tentativas = 0
    print('>>> Aguardando ECAC')
    while not _find_img('gerar_relatorio.png', conf=0.9):
        time.sleep(1)
        timer += 1
        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('erro_sistema_3.png', conf=0.9):
            return '❗ Solicitação rejeitada pelo sistema, tente novamente mais tarde.'
            
        # se aparecer a tela de consulta em processamento, retorna o erro
        if _find_img('erro_sistema_7.png', conf=0.9):
            return '❗ Consulta não liberada pelo sistema, tente mais tarde.'
            
        # se aparecer a tela de mensagens do ecac, retorna o erro
        if _find_img('mensagens_importantes_ecac.png', conf=0.9):
            return '❗ Este CPF possuí mensagens importantes no ECAC, não é possível emitir o relatório até que a mensagem seja aberta.'
         
        if _find_img('erro_sistema_5.png', conf=0.9):
            reiniciar_ecac()
            timer = 0
            tentativas += 1
            
        if _find_img('erro_sistema_8.png', conf=0.9):
            reiniciar_ecac()
            timer = 0
            tentativas += 1
            
        if _find_img('erro_sistema_9.png', conf=0.9):
            reiniciar_ecac()
            timer = 0
            tentativas += 1
        
        # se aparecer a tela de consulta em processamento, retorna o erro
        if _find_img('pagina_dirf.png', conf=0.9):
            reiniciar_ecac()
            timer = 0
            tentativas += 1
            
        # se demorar 5 segundos para o botão de emissão da certidão aparecer, verifica se a tela de login do ecac stá bugada
        # se a tela de login do ecac estiver bugada, fecha a janela, recarrega a página no conferir e clica no botão de CND do ecac novamente
        if timer > 60:
            reiniciar_ecac()
            timer = 0
            tentativas += 1

        if tentativas > 5:
            # se demorar muito pata carregar, retorna o erro
            return '❌ Erro ao acessar o ECAC, sistema demorou muito para responder.'

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
                return '✔ OK'
        if timer > 60:
            print('>>> Erro no ECAC, tentando novamente')
            return '✔ OK'
        time.sleep(1)
        timer += 1
        
    # clica no botão para emitir o relatório
    _click_img('botao_gerar_relatorio.png', conf=0.9)

    return '✔ OK'


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
    
    while True:
        try:
            pyperclip.copy(pasta_final)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass

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

    return '✔ Relatório emitido com sucesso'


def verifica_relatorio(pasta_analise, pasta_final, pasta_final_sem_pendencias):
    print('>>> Analisando relatório')
    # Analisa cada pdf que estiver na pasta
    for arquivo in os.listdir(pasta_analise):
        print(f'Arquivo: {arquivo}')
        # Abrir o pdf
        arq = os.path.join(pasta_analise, arquivo)
    
        with fitz.open(arq) as pdf:
            # Para cada página do pdf
            situacao_1 = ''
            situacao_2 = ''
            for count, page in enumerate(pdf):
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                
                if re.compile(r'Diagnóstico Fiscal na Receita Federal e Procuradoria-Geral da Fazenda Nacional.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                    os.makedirs(pasta_final_sem_pendencias, exist_ok=True)
                    situacao = '✔ Sem pendências'
                    pasta_final = os.path.join(pasta_final_sem_pendencias, arquivo.replace('.pdf', '-Sem pendências.pdf'))
                    break
                else:
                    os.makedirs(pasta_final, exist_ok=True)
                    if re.compile(r'Diagnóstico Fiscal na Receita Federal').search(textinho):
                        if not re.compile(r'Diagnóstico Fiscal na Receita Federal.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                            situacao_1 = 'Pendência na Receita Federal'
                        else:
                            situacao_1 = 'Sem pendência na Receita Federal'
                            
                    if re.compile(r'Diagnóstico Fiscal na Procuradoria-Geral da Fazenda Nacional').search(textinho):
                        if not re.compile(r'Diagnóstico Fiscal na Procuradoria-Geral da Fazenda Nacional.+\nNão foram detectadas pendências/exigibilidades').search(textinho):
                            situacao_2 = 'Pendência na Procuradoria Geral da Fazenda Nacional'
                        else:
                            situacao_2 = 'Sem pendencias na Procuradoria Geral da Fazenda Nacional'
                    
                    situacao = f'❗ Com pendências;{situacao_1};{situacao_2}'
                    
        shutil.move(arq, pasta_final)
    return situacao


@_time_execution
def run():
    pasta_analise = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\ignore\Relatórios'
    pasta_final = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\execução\Relatórios'
    pasta_final_sem_pendencias = r'V:\Setor Robô\Scripts Python\Ecac\Relatórios Fiscais Conferir\execução\Relatórios Sem Pendências'
    
    os.makedirs(pasta_analise, exist_ok=True)
    
    # limpa a pasta de download
    for file in os.listdir(pasta_analise):
        os.remove(os.path.join(pasta_analise, file))
        
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
        
        print('>>> Abrindo o site do Conferir')
        _abrir_chrome('https://portal.conferironline.com.br/dashboard')
        resultado = busca_cpf(cpf)
        if resultado == '✔ OK':
            resultado = consulta_relatorio()
            
        if resultado == '✔ OK':
            resultado = salvar_pdf(pasta_analise)
            situacao = verifica_relatorio(pasta_analise, pasta_final, pasta_final_sem_pendencias)
        else:
            situacao = ''

        p.hotkey('ctrl', 'w')
        print(f'{resultado} - {situacao}'.replace(';', ' - '))
        _escreve_relatorio_csv(f'{cpf};{nome};{resultado[2:]};{situacao[2:]}', nome='Relatórios Fiscais Conferir')


if __name__ == '__main__':
    run()
