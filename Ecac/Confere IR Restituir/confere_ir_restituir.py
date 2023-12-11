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
    time.sleep(1)
    

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
    time.sleep(1)
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


def abre_irpf_atual():
    # aguarda o menu do ecac aparecer
    while not _find_img('acoes_ecac.png', conf=0.9):
        time.sleep(1)
    
    # clica no menu do ecac
    _click_img('acoes_ecac.png', conf=0.9)
    time.sleep(1)
    
    # aguarda o botão de CND no ecac aparecer, se aparecer uma mensagem sobre login inválido retorna o erro
    while not _find_img('meu_irpf_ecac.png', conf=0.9):
        if _find_img('login_invalido.png', conf=0.9):
            return '❗ Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.'
        if _find_img('meu_irpf_ecac_2.png', conf=0.9):
            break
        p.press('PgDn')
        time.sleep(1)
    
    print('>>> Acessando ECAC')
    # clica no botão do Meu IRPF do ECAC
    _click_position_img('meu_irpf_ecac.png', '+', pixels_y=92, conf=0.9)
    _click_position_img('meu_irpf_ecac_2.png', '+', pixels_y=50, conf=0.9)
    
    timer = 0
    tentativas = 0
    # aguarda a tela do Meu IRPF
    print('>>> Aguardando ECAC')
    while not _find_img('irpf_atual.png', conf=0.9):
        time.sleep(1)
        timer += 1
        # se a tela de login do ecac estiver bugada, fecha a janela, recarrega a página no conferir e clica no botão de CND do ecac novamente
        if timer > 60:
            print('>>> Erro no ECAC, tentando novamente')
            reiniciar_ecac()
            timer = 0
            
            tentativas += 1
            if tentativas >= 5:
                return '❌ Não é possível acessar o ECAC, sistema demorou muito pra responder'

    time.sleep(2)
    # clica para abrir a tela do relatório
    print('>>> Consultando IRPF')
    _click_img('irpf_atual.png', conf=0.9)
    
    # aguarda a tela referênte ao IRPF atual
    while not _find_img('exercicio.png', conf=0.9):
        time.sleep(1)
    
    return '✔ OK'


def confere_restituir(cpf, nome, pasta_restituir, pasta_pagar):
    time.sleep(1)
    timer = 0
    while True:
        print('>>> Verificando se existe restituição')
        pendencias = ''
        pendencias_pdf = ''
        if _find_img('pendencias.png', conf=0.9):
            pendencias = 'Com Pendências'
            pendencias_pdf = ' - Com Pendências'
        
        if _find_img('dependente.png', conf=0.9):
            return '❗ Contribuinte consta como dependente em outra declaração.;' + pendencias
        
        if _find_img('nao_entregue.png', conf=0.9):
            return '❌ Não consta entrega de declaração para este ano.;' + pendencias
        
        if _find_img('declaracao_omissa.png', conf=0.9):
            return '❌❗ Omisso. Contribuinte obrigado à entrega de Declaração IRPF.;' + pendencias
        
        # sem saldo de imposto
        if _find_img('sem_saldo.png', conf=0.9):
            return '✔ Sem saldo de imposto;0,0;' + pendencias
        
        # imposto a restituir
        if _find_img('restituir.png', conf=0.9):
            salvar_pdf(f'{cpf} - {nome} - Imposto a Restituir{pendencias_pdf}', pasta_restituir)
            valor = copia_valor()
            return f'✔ Imposto a restituir;{valor};' + pendencias
        
        # imposto a pagar
        if _find_img('pagar.png', conf=0.9):
            '''salvar_pdf(f'{cpf} - {nome} - Imposto a Restituir{pendencias_pdf}', pasta_pagar)'''
            valor = copia_valor()
            return f'❗ Imposto a pagar;{valor};' + pendencias
        
        p.press('pgdn')
        timer += 1
        if timer > 60:
            return 'erro'
      

def confere_guias(cpf, nome, pasta_guias):
    print('>>> Verificando se existe restituição')
    
    
    
    
    return f'✔ Guia gerada'


def reiniciar_ecac():
    print('>>> Erro no ECAC, tentando novamente')
    p.click(1000, 10)
    time.sleep(1)
    p.hotkey('ctrl', 'w')
    time.sleep(1)
    p.press('f5')
    time.sleep(2)
    
    # aguarda o botão de CND no ecac aparecer, se aparecer uma mensagem sobre login inválido retorna o erro
    while not _find_img('meu_irpf_ecac.png', conf=0.9):
        if _find_img('login_invalido.png', conf=0.9):
            return '❗ Os serviços só estão disponíveis caso o login esteja válido! Verifique na aba ECAC o login e senha por favor.'
        if _find_img('meu_irpf_ecac_2.png', conf=0.9):
            break
        p.press('PgDn')
        time.sleep(1)
    
    print('>>> Acessando ECAC')
    # clica no botão do Meu IRPF do ECAC
    _click_position_img('meu_irpf_ecac.png', '+', pixels_y=92, conf=0.9)
    _click_position_img('meu_irpf_ecac_2.png', '+', pixels_y=50, conf=0.9)
 
    
def salvar_pdf(arquivo, pasta_final):
    p.hotkey('ctrl', 'p')
    
    print('>>> Aguardando tela de impressão')
    _wait_img('tela_imprimir.png', conf=0.9)
    
    print('>>> Salvando PDF')
    imagens = ['print_to_pdf.png', 'print_to_pdf_2.png']
    for img in imagens:
        # se não estiver selecionado para salvar como PDF, seleciona para salvar como PDF
        if _find_img(img, conf=0.9) or _find_img(img, conf=0.9):
            _click_img(img, conf=0.9)
            # aguarda aparecer a opção de salvar como PDF e clica nela
            _wait_img('salvar_como_pdf.png', conf=0.9)
            _click_img('salvar_como_pdf.png', conf=0.9)
    
    # aguarda aparecer o botão de salvar e clica nele
    while not _find_img('botao_salvar.png', conf=0.9):
        time.sleep(1)
    
    p.click(p.locateCenterOnScreen(f'imgs/botao_salvar.png', confidence=0.9))
    
    while not _find_img('salvar_como.png', conf=0.9):
        print('>>> Salvando relatório')
        p.click(p.locateCenterOnScreen(f'imgs/botao_salvar.png', confidence=0.9))
        time.sleep(1)
    
    time.sleep(1)
    os.makedirs(pasta_final, exist_ok=True)
    
    while True:
        try:
            pyperclip.copy(arquivo)
            time.sleep(1)
            pyperclip.copy(arquivo)
            time.sleep(1)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    
    time.sleep(1)
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


def copia_valor():
    _click_position_img('valor.png', '+', pixels_x=23, conf=0.9, clicks=2)
    # copia o valor
    while True:
        try:
            time.sleep(1)
            p.hotkey('ctrl', 'c')
            time.sleep(1)
            p.hotkey('ctrl', 'c')
            time.sleep(1)
            valor = pyperclip.paste()
            return valor
        except:
            pass


@_time_execution
def run():
    consulta = p.confirm(buttons=['Consulta IR a restituir', 'Consulta e gera DARF'])
    
    if consulta == 'Consulta IR a restituir':
        pasta_restituir = r'V:\Setor Robô\Scripts Python\Ecac\Confere IR Restituir\Execução\Imposto a Restituir'
        pasta_pagar = r'V:\Setor Robô\Scripts Python\Ecac\Confere IR Restituir\Execução'
    if consulta == 'Consulta e gera DARF':
        pasta_guias = r'V:\Setor Robô\Scripts Python\Ecac\Confere IR Restituir\Execução\Guias DARF'
    
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
        
        while True:
            print('>>> Abrindo o site do Conferir')
            time.sleep(1)
            _abrir_chrome('https://portal.conferironline.com.br', fechar_janela=False)
            resultado = busca_cpf(cpf)
            if resultado == '✔ OK':
                resultado = abre_irpf_atual()
            
            if resultado == '✔ OK':
                if consulta == 'Consulta IR a restituir':
                    situacao = confere_restituir(cpf, nome, pasta_restituir, pasta_pagar)
                if consulta == 'Consulta e gera DARF':
                    situacao = confere_guias(cpf, nome, pasta_guias)
                    
                if situacao != 'erro':
                    p.hotkey('ctrl', 'w')
                    break
            else:
                situacao = '✖ '
                p.hotkey('ctrl', 'w')
                break
            
            p.hotkey('ctrl', 'w')
            
        p.hotkey('ctrl', 'w')
        print(f'{resultado} - {situacao}'.replace(';', ' - '))
        _escreve_relatorio_csv(f'{cpf};{nome};{resultado[2:]};{situacao[2:]}', nome='Confere Imposto a Restituir')


if __name__ == '__main__':
    run()
