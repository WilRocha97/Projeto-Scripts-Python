import pyautogui as p
import time
import os
import pyperclip
from sys import path

path.append(r'..\..\_comum')
from sn_comum import _new_session_sn_driver
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def verificacoes(empresa, andamentos):
    cnpj, comp, ano, venc = empresa
    if _find_img('NaoOptante.png'):
        _escreve_relatorio_csv(f'{cnpj};Não optante no ano atual;{comp};{ano}', nome=andamentos)
        print('❌ Não optante no ano atual')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('JaExistePag.png'):
        _escreve_relatorio_csv(f'{cnpj};Já existe pagamento para essa competência;{comp};{ano}', nome=andamentos)
        print('❌ Já existe pagamento para essa competência')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('NaoTemAnterior.png'):
        _escreve_relatorio_csv(f'{cnpj};Não consta pagamento de anos anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Não consta pagamento de anos anteriores')
        p.hotkey('ctrl', 'w')
        return False
    return True


def imprimir(empresa, andamentos):
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(r'{}\{}'.format(e_dir, 'Boletos'), exist_ok=True)

    cnpj, comp, ano, venc = empresa
    venc = venc.replace('/','-')
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
    _escreve_relatorio_csv(f'{cnpj};Guia gerada;{comp};{ano}', nome=andamentos)
    print('✔ Guia gerada')
    time.sleep(8)
    p.hotkey('ctrl', 'w')


def boleto_mei(empresa, andamentos):
    cnpj, comp, ano, venc = empresa
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

    link = 'http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao'
    
    _click_img('Maxi.png', conf=0.9, timeout=1)
    p.click(1150, 51)
    time.sleep(1)
    p.write(link.lower())
    time.sleep(1)
    p.press('enter')
    time.sleep(3)
    
    # espera o site abrir
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
        _escreve_relatorio_csv(f'{cnpj};Não optante;{comp};{ano}', nome=andamentos)
        print('❌ Não optante')
        p.hotkey('ctrl', 'w')
        return False

    if not _find_img('2023.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Ano não liberado;{comp};{ano}', nome=andamentos)
        print('❌ Ano não liberado')
        p.hotkey('ctrl', 'w')
        return False

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
        _escreve_relatorio_csv(f'{cnpj};Não consta pagamento de anos anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Não consta pagamento de anos anteriores')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('NaoTemAno.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Não optante;{comp};{ano}', nome=andamentos)
        print('❌ Não optante')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('AntesDeProsseguir.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores')
        p.hotkey('ctrl', 'w')
        return False

    # Esperar abrir a tela das guias
    _wait_img('Selecione.png', conf=0.9, timeout=-1)
    time.sleep(3)

    if not _find_img('mes' + comp + '.png', conf=0.99):
        p.press('pgDn')
        time.sleep(5)
        if not _find_img('mes' + comp + '.png', conf=0.99):
            _escreve_relatorio_csv(f'{cnpj};Competência não liberada;{comp};{ano}', nome=andamentos)
            print('❌ Competência não liberada')
            p.hotkey('ctrl', 'w')
            return False

    _click_img('mes' + comp + '.png', conf=0.99)

    # Descer a página e clicar em apurar
    time.sleep(1)
    p.press('pgDn')
    time.sleep(1)
    _click_img('Apurar.png', conf=0.9)
    time.sleep(2)

    if _find_img('MesBaixado.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Competência não habilitada;{comp};{ano}', nome=andamentos)
        print('❌ Competência não habilitada')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('JaConsta.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Já consta pagamento;{comp};{ano}', nome=andamentos)
        print('❗ Já consta pagamento')
        p.hotkey('ctrl', 'w')
        return False

    _wait_img('Voltar.png', conf=0.9, timeout=-1)
    return True
    

@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        cnpj, comp, ano, venc = empresa
        andamentos = f'Boletos MEI {comp}-{ano}'
        
        if not boleto_mei(empresa, andamentos):
            continue
            
        if not verificacoes(empresa, andamentos):
            continue
            
        # Salvar a guia
        imprimir(empresa, andamentos)
    
    time.sleep(2)
    p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    run()
