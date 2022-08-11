# -*- coding: utf-8 -*-

import pyperclip, time, os, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def salvar(cnpj):
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(r'{}\{}'.format(e_dir, 'Recibos'), exist_ok=True)
    # Esperar o PDF gerar
    time.sleep(1)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy(r'V:\Setor Robô\Scripts Python\Ecac\recibos dctf (tapa buraco)\{}\Recibos'.format(e_dir))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')

    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('SalvarComo.png', 0.9):
        p.press('s')
        time.sleep(1)

    print('✔ Recibo salvo')


def abrir():
    # Abrir o site
    if _find_img('Chrome.png', conf=0.9):
        _click_img('Chrome.png', conf=0.9)
    else:
        p.hotkey('win', 'm')
        time.sleep(0.5)
        _click_img('ChromeAtalho.png', conf=0.9, clicks=2)

    link = 'https://cav.receita.fazenda.gov.br/autenticacao/login'
    while not _find_img('link.png', conf=0.9, ):
        time.sleep(5)
        p.moveTo(1163, 377)
        p.click()
    _click_img('Maxi.png', conf=0.9, timeout=2)
    _click_img('link.png', conf=0.9)
    p.write(link)
    time.sleep(1)
    p.press('enter')
    time.sleep(3)

    if not _find_img('AlterarEmpresa.png', conf=0.9):
        _wait_img('GovBR.png', conf=0.9, timeout=-1)
        _click_img('GovBR.png', conf=0.9)

        _wait_img('X.png', conf=0.9, timeout=-1)
        _click_img('X.png', conf=0.9)

        _wait_img('Certificado.png', conf=0.9, timeout=-1)
        _click_img('Certificado.png', conf=0.9, )
        time.sleep(2)
        p.press('enter')
        _wait_img('AlterarEmpresa.png', conf=0.9, timeout=-1)


def recibos(index, empresas, data_inicial, data_final):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        cnpj, cert = empresa

        _indice(count, total_empresas, empresa)
        p.hotkey('win', 'm')

        if not _find_img('AlterarEmpresa.png', conf=0.9):
            abrir()

        while _find_img('Mensagens.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            abrir()

        _wait_img('AlterarEmpresa.png', conf=0.9, timeout=-1)
        _click_img('AlterarEmpresa.png', conf=0.9, )

        _wait_img('CNPJ.png', conf=0.9, timeout=-1)
        _click_img('CNPJ.png', conf=0.9)
        time.sleep(1)
        p.write(cnpj)
        time.sleep(1)

        _click_img('Alterar.png', conf=0.9)
        time.sleep(3)

        if _find_img('Mensagens.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            _escreve_relatorio_csv('{};Mensagens no ECAC'.format(cnpj))
            print('❌ Mensagens no ECAC')
            continue

        if _find_img('atencao.png', conf=0.9):
            _escreve_relatorio_csv('{};Erro no login'.format(cnpj))
            print('❌ Erro no login')
            continue

        if not _find_img('SemAcesso.png', conf=0.9):
            _escreve_relatorio_csv('{};CPNJ sem acesso ao ECAC'.format(cnpj))
            print('❌ CPNJ sem acesso ao ECAC')
            continue

        link = 'https://dctfweb.cav.receita.fazenda.gov.br/aplicacoesweb/DCTFWeb/Default.aspx'
        _click_img('link.png', conf=0.9)
        p.write(link)
        time.sleep(1)
        p.press('enter')

        _wait_img('DCTFweb.png', conf=0.9, timeout=-1)

        _click_img('DataInicial.png', conf=0.9)
        time.sleep(0.5)
        p.write(data_inicial)
        time.sleep(0.5)

        _click_img('DataFinal.png', conf=0.9)
        time.sleep(0.5)
        p.write(data_final)
        time.sleep(0.5)

        _click_img('Pesquisar.png', conf=0.9)
        time.sleep(5)

        indice = 0
        if _find_img('EmAndamento.png', conf=0.9):
            _escreve_relatorio_csv('{};Existe declaração em andamento'.format(cnpj))
            print('❗ Existe declaração em andamento')

        elif _find_img('ReciboAno.png', conf=0.9):
            _click_img('ReciboAno.png', conf=0.9)
            time.sleep(1)
            _click_img('DCTFweb.png', conf=0.9)
            time.sleep(0.5)

            _click_img('ReciboMes.png', conf=0.9)
            time.sleep(1)
            _click_img('DCTFweb.png', conf=0.9)
            indice = + 2

        elif _find_img('Recibo.png', conf=0.9):
            _click_img('Recibo.png', conf=0.9)
            time.sleep(1)
            _click_img('DCTFweb.png', conf=0.9)
            indice = + 1

        else:
            _escreve_relatorio_csv('{};Nenhum recibo encontrado'.format(cnpj))
            print('❌ Nenhum recibo encontrado')

        if indice == 1:
            _escreve_relatorio_csv('{};Salvou recibo'.format(cnpj))
            print('✔ Recibo salvo')
        if indice == 2:
            _escreve_relatorio_csv('{};Salvou recibo anual e mensal'.format(cnpj))
            print('✔ Recibo anual e mensal salvos')


@_time_execution
def run():
    data_inicial = p.prompt(text='Qual a data inicial da apuração?', title='Script incrível', default='00/00/0000')
    data_final = p.prompt(text='Qual a data final da apuração?', title='Script incrível', default='00/00/0000')
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    recibos(index, empresas, data_inicial, data_final)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
