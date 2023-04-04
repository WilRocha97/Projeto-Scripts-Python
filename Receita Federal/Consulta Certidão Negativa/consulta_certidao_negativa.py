import pyautogui as p
import time
import os
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def verificacoes(consulta_tipo, andamento, empresa):
    identificacao, nome = empresa
    if _find_img('inscricao_cancelada.png', conf=0.9):
        _escreve_relatorio_csv('{};{};inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB'.format(identificacao, nome), nome=andamento)
        print('❌ inscrição cancelada de ofício pela Secretaria Especial da Receita Federal do Brasil - RFB')
        return False

    if _find_img('NaoFoiPossivel.png', conf=0.9):
        _escreve_relatorio_csv('{};{};Não foi possível concluir a consulta'.format(identificacao, nome), nome=andamento)
        print('❌ Não foi possível concluir a consulta')
        return False

    if _find_img('InfoInsuficiente.png', conf=0.9):
        _escreve_relatorio_csv('{};{};As informações disponíveis na Secretaria da Receita Federal do Brasil - RFB sobre o contribuinte '
                               'são insuficientes para a emissão de certidão por meio da Internet.'.format(identificacao, nome), nome=andamento)
        print('❌ Informações insuficientes para a emissão de certidão online')
        return False

    if _find_img('Matriz.png', conf=0.9):
        _escreve_relatorio_csv('{};{};A certidão deve ser emitida para o {} da matriz'.format(identificacao, nome, consulta_tipo), nome=andamento)
        print(f'❌ A certidão deve ser emitida para o {consulta_tipo} da matriz')
        return False
    
    if _find_img('cpf_invalido.png', conf=0.9):
        _escreve_relatorio_csv('{};{};CPF inválido'.format(identificacao, nome), nome=andamento)
        print(f'❌ CPF inválido')
        return False
    
    if _find_img('ErroNaConsulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)
        p.press('enter')

    return True


def salvar(consulta_tipo, andamento, empresa, pasta):
    identificacao, nome = empresa
    # espera abrir a tela de salvar o arquivo
    while not _find_img('SalvarComo.png', conf=0.9):
        time.sleep(1)
        if _find_img('em_processamento.png', conf=0.9):
            consulta(consulta_tipo, empresa)

        if _find_img('NovaCertidao.png', conf=0.9):
            _click_img('NovaCertidao.png', conf=0.9)

        if not verificacoes(consulta_tipo, andamento, empresa):
            p.hotkey('ctrl', 'w')
            return False

    # escreve o nome do arquivo (.upper() serve para deixar em letra maiúscula)
    time.sleep(1)
    p.write(f'{nome.upper()} - {identificacao} - Certidao')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Receita Federal\Consulta Certidão Negativa\{}\{}'.format(e_dir, pasta))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')

    # salvar o arquivo
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(2)
    
    if _find_img('substituir.png', conf=0.9):
        p.hotkey('alt', 's')
        
    return True


def consulta(consulta_tipo, empresa):
    identificacao, nome = empresa
    p.hotkey('win', 'm')

    # Abrir o site
    if _find_img('Chrome.png', conf=0.99):
        pass
    elif _find_img('ChromeAberto.png', conf=0.99):
        _click_img('ChromeAberto.png', conf=0.99)
    else:
        time.sleep(0.5)
        _click_img('ChromeAtalho.png', conf=0.9, clicks=2)
        while not _find_img('Google.png', conf=0.9, ):
            time.sleep(5)
            p.moveTo(1163, 377)
            p.click()

    if consulta_tipo == 'CNPJ':
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir'
    else:
        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PF/Emitir'

    _click_img('Maxi.png', conf=0.9, timeout=1)
    p.click(1150, 51)
    time.sleep(1)
    p.write(link.lower())
    time.sleep(1)
    p.press('enter')
    time.sleep(3)

    # espera o site abrir
    while not _wait_img(f'Informe{consulta_tipo}.png', conf=0.9, timeout=-1):
        time.sleep(1)

    _click_img(f'Informe{consulta_tipo}.png', conf=0.9)
    time.sleep(1)

    p.write(identificacao)
    time.sleep(1)
    p.press('enter')

    return True


@_time_execution
def run():
    consulta_tipo = p.confirm(text='Qual o tipo da consulta?', buttons=('CNPJ', 'CPF'))
    pasta = f'Certidões {consulta_tipo}'
    os.makedirs(r'{}\{}'.format(e_dir, pasta), exist_ok=True)
    andamento = f'Certidão Negativa {consulta_tipo}'

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        try:
            consulta(consulta_tipo, empresa)
            if not salvar(consulta_tipo, andamento, empresa, pasta):
                p.hotkey('ctrl', 'w')
                continue
            
            identificacao, nome = empresa
            print('✔ Certidão gerada')
            _escreve_relatorio_csv('{};{};{} gerada'.format(identificacao, nome, andamento), nome=andamento)
            time.sleep(1)

        except:
            time.sleep(2)
            p.hotkey('ctrl', 'w')

    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
