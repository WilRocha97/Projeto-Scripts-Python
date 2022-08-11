import pyautogui as p
import time
import os
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from sn_comum import _new_session_sn_driver
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def verificacoes(empresa):
    cnpj, comp, ano, venc = empresa
    if _find_img('NaoOptante.png'):
        _escreve_relatorio_csv(';'.join([cnpj, 'Não optante no ano atual', comp, ano]), nome='Boletos MEI')
        print('❌ Não optante no ano atual')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('JaExistePag.png'):
        _escreve_relatorio_csv(';'.join([cnpj, 'Já existe pagamento para essa competência', comp, ano]), nome='Boletos MEI')
        print('❌ Já existe pagamento para essa competência')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('NaoTemAnterior.png'):
        _escreve_relatorio_csv(';'.join([cnpj, 'Não consta pagamento de anos anteriores', comp, ano]), nome='Boletos MEI')
        print('❌ Não consta pagamento de anos anteriores')
        p.hotkey('ctrl', 'w')
        return False
    return True


def imprimir(empresa):
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(r'{}\{}'.format(e_dir, 'Boletos'), exist_ok=True)

    cnpj, comp, ano, venc = empresa
    p.moveTo(647, 468)
    p.click()
    _wait_img('SalvarComo.png', conf=0.9, timeout=-1)
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    p.write(cnpj + ';DAS;' + comp + ';' + ano + ';' + venc + ';Guia do MEI ' + comp + '-' + ano)
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Portal SN\Boletos MEI\{}\{}'.format(e_dir, 'Boletos'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('Substituir.png'):
        p.press('s')
    time.sleep(1)
    _escreve_relatorio_csv(';'.join([cnpj, 'Guia gerada', comp, ano]), nome='Boletos MEI')
    print('✔ Guia gerada')
    time.sleep(5)
    p.hotkey('ctrl', 'w')


def boleto_mei(empresas, index):
    total_empresas = empresas[index:]

    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)

        cnpj, comp, ano, venc = empresa

        # Abrir o site
        if _find_img('Chrome.png', conf=0.9):
            _click_img('Chrome.png', conf=0.9)
        else:
            p.hotkey('win', 'm')
            time.sleep(0.5)
            _click_img('ChromeAtalho.png', conf=0.9, clicks=2)

        link = 'http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao'
        
        while not _find_img('link.png', conf=0.9,):
            time.sleep(5)
            p.moveTo(1163, 377)
            p.click()
        _click_img('Maxi.png', conf=0.9, timeout=2)
        _click_img('link.png', conf=0.9)
        p.write(link)
        time.sleep(1)
        p.press('enter')
        time.sleep(3)

        while not _wait_img('InformeCNPJ.png', conf=0.9, timeout=-1):
            time.sleep(1)
        # Fazer login com CNPJ
        _click_img('CNPJ.png', conf=0.9)
        p.write(cnpj)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(0.5)

        # Emitir guia de pagamento DAS
        _wait_img('EmitirGuia.png', conf=0.9, timeout=-1)
        _click_img('EmitirGuia.png', conf=0.9)

        # Selecionar o ano e clicar em ok
        _wait_img('Ano.png', conf=0.95, timeout=-1)
        time.sleep(1)

        _click_img('Ano.png', conf=0.95)
        time.sleep(1)

        if _find_img('NaoOptante2.png', conf=0.95):
            _escreve_relatorio_csv(';'.join([cnpj, 'Não optante', comp, ano]), nome='Boletos MEI')
            print('❌ Não optante\n')
            p.hotkey('ctrl', 'w')
            continue

        if not _find_img('2022.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cnpj, 'Ano não liberado', comp, ano]), nome='Boletos MEI')
            print('❌ Ano não liberado\n')
            p.hotkey('ctrl', 'w')
            continue

        # while not _find_img('AnoDeConsulta.png', conf=0.9):
        p.press('up')
        time.sleep(0.5)
        p.press('up')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)
        _click_img('ok.png', conf=0.9)
        time.sleep(1)

        if _find_img('NaoTemAnterior.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cnpj, 'Não consta pagamento de anos anteriores', comp, ano]), nome='Boletos MEI')
            print('❌ Não consta pagamento de anos anteriores\n')
            p.hotkey('ctrl', 'w')
            continue

        if _find_img('NaoTemAno.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cnpj, 'Não optante', comp, ano]), nome='Boletos MEI')
            print('❌ Não optante\n')
            p.hotkey('ctrl', 'w')
            continue

        if _find_img('AntesDeProsseguir.png', conf=0.9):
            _escreve_relatorio_csv(';'.join(
                [cnpj, 'Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores', comp, ano]), nome='Boletos MEI')
            print('❌ Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores\n')
            p.hotkey('ctrl', 'w')
            continue

        # Esperar abrir a tela das guias
        _wait_img('Selecione.png', conf=0.9, timeout=-1)
        time.sleep(3)

        if not _find_img('mes' + comp + '.png', conf=0.99):
            p.press('pgDn')
            time.sleep(5)
            if not _find_img('mes' + comp + '.png', conf=0.99):
                _escreve_relatorio_csv(';'.join([cnpj, 'Competência não liberada', comp, ano]), nome='Boletos MEI')
                print('❌ Competência não liberada\n')
                p.hotkey('ctrl', 'w')
                continue

        _click_img('mes' + comp + '.png', conf=0.99)

        # Descer a página e clicar em apurar
        time.sleep(1)
        p.press('pgDn')
        time.sleep(1)
        _click_img('Apurar.png', conf=0.9)
        time.sleep(2)

        if _find_img('MesBaixado.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cnpj, 'Competência não habilitada', comp, ano]), nome='Boletos MEI')
            print('❌ Competência não habilitada\n')
            p.hotkey('ctrl', 'w')
            continue

        if _find_img('JaConsta.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cnpj, 'Já consta pagamento', comp, ano]), nome='Boletos MEI')
            print('❗ Já consta pagamento\n')
            p.hotkey('ctrl', 'w')
            continue

        _wait_img('Voltar.png', conf=0.9, timeout=-1)

        if not verificacoes(empresa):
            continue

        # Salvar a guia
        imprimir(empresa)


@_time_execution
def run():
    os.makedirs('execucao/Boletos', exist_ok=True)

    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    # try:
    boleto_mei(empresas, index)
    # except:
    time.sleep(2)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    # _new_session_sn_driver('03684902000182')   Não funciona porque o site reconhece o robô
    run()
