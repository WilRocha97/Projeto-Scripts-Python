import pyautogui as p
import time
import os
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def verificacoes(andamento, empresa):
    cnpj, nome = empresa
    if _find_img('NaoFoiPossivel.png', conf=0.9):
        _escreve_relatorio_csv('{};{};Não foi possível concluir a consulta'.format(cnpj, nome), nome=andamento)
        print('❌ Não foi possível concluir a consulta')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('InfoInsuficiente.png', conf=0.9):
        _escreve_relatorio_csv('{};{};As informações disponíveis na Secretaria da Receita Federal do Brasil - RFB sobre o contribuinte '
                               'são insuficientes para a emissão de certidão por meio da Internet.'.format(cnpj, nome), nome=andamento)
        print('❌ Informações insuficientes para a emissão de certidão online')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('Matriz.png', conf=0.9):
        _escreve_relatorio_csv('{};{};A certidão deve ser emitida para o CNPJ da matriz'.format(cnpj, nome), nome=andamento)
        print('❌ A certidão deve ser emitida para o CNPJ da matriz')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('ErroNaConsulta.png', conf=0.9):
        p.press('enter')
        time.sleep(2)
        p.press('enter')

    return True


def salvar(andamento, empresa):
    cnpj, nome = empresa
    # espera abrir a tela de salvar o arquivo
    while not _find_img('SalvarComo.png', conf=0.9):
        time.sleep(1)
        if _find_img('NovaCertidao.png', conf=0.9):
            _click_img('NovaCertidao.png', conf=0.9)

        if not verificacoes(andamento, empresa):
            return False

    # escreve o nome do arquivo (.upper() serve para deixar em letra maiuscula)
    time.sleep(1)
    p.write(f'{nome.upper()} - {cnpj} - Certidao')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Receita Federal\Consulta Certidão Negativa\{}\{}'.format(e_dir, 'Certidões'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')

    # salvar o arquivo
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(2)
    return True


def consulta(andamento, empresas, index):
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)

        cnpj, nome = empresa

        p.hotkey('win', 'm')

        # Abrir o site
        if _find_img('Chrome.png', conf=0.99):
            pass
        elif _find_img('ChromeAberto.png', conf=0.99):
            _click_img('ChromeAberto.png', conf=0.99)
        else:
            p.hotkey('win', 'm')
            time.sleep(0.5)
            _click_img('ChromeAtalho.png', conf=0.9, clicks=2)
            while not _find_img('Google.png', conf=0.9, ):
                time.sleep(5)
                p.moveTo(1163, 377)
                p.click()

        link = 'https://solucoes.receita.fazenda.gov.br/Servicos/certidaointernet/PJ/Emitir'

        _click_img('Maxi.png', conf=0.9, timeout=2)
        p.click(1208, 51)
        time.sleep(1)
        p.write(link.lower())
        time.sleep(1)
        p.press('enter')
        time.sleep(3)

        # espera o site abrir
        while not _wait_img('InformeCNPJ.png', conf=0.9, timeout=-1):
            time.sleep(1)

        _click_img('InformeCNPJ.png', conf=0.9)
        time.sleep(1)

        p.write(cnpj)
        time.sleep(1)
        p.press('enter')

        if not salvar(andamento, empresa):
            continue

        print('✔ Certidão gerada')
        _escreve_relatorio_csv('{};{};{} gerada'.format(cnpj, nome, andamento), nome=andamento)
        time.sleep(1)

    p.hotkey('ctrl', 'w')


@_time_execution
def run():
    os.makedirs(r'{}\{}'.format(e_dir, 'Certidões'), exist_ok=True)
    andamento = 'Certidão Negativa'

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    try:
        consulta(andamento, empresas, index)
    except:
        time.sleep(2)
        p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
