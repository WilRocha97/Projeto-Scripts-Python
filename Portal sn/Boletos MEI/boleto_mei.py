import pyautogui as p
import time
import os
import pyperclip

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _abrir_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start


def verificacoes(empresa, andamentos):
    cnpj, comp, ano, venc = empresa
    if _find_img('nao_optante.png'):
        _escreve_relatorio_csv(f'{cnpj};Não optante no ano atual;{comp};{ano}', nome=andamentos)
        print('❌ Não optante no ano atual')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('ja_existe_pag.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Já existe pagamento para essa competência;{comp};{ano}', nome=andamentos)
        print('❌ Já existe pagamento para essa competência')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('nao_tem_anterior.png'):
        _escreve_relatorio_csv(f'{cnpj};Não consta pagamento de anos anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Não consta pagamento de anos anteriores')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('debito_automatico.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};O PA está em débito automático;{comp};{ano}', nome=andamentos)
        print('❕ O PA está em débito automático')
        p.hotkey('ctrl', 'w')
        return False
    if _find_img('nao_tem_das.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Não há DAS a ser emitido;{comp};{ano}', nome=andamentos)
        print('❕ Não há DAS a ser emitido')
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
    _wait_img('salvar_como.png', conf=0.9, timeout=-1)
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
    if _find_img('substituir.png'):
        p.press('s')
    time.sleep(1)
    _escreve_relatorio_csv(f'{cnpj};Guia gerada;{comp};{ano}', nome=andamentos)
    print('✔ Guia gerada')
    time.sleep(8)
    p.hotkey('ctrl', 'w')


def boleto_mei(empresa, andamentos):
    cnpj, comp, ano, venc = empresa
    
    # espera o site abrir
    while not _wait_img('informe_cnpj.png', conf=0.9, timeout=-1):
        time.sleep(1)
    
    print('>>> inserindo CNPJ')
    # Fazer login com CNPJ
    _click_img('cnpj.png', conf=0.9)
    p.write(cnpj)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    # Emitir guia de pagamento DAS
    while not _find_img('emitir_guia.png', conf=0.9):
        time.sleep(1)
        if _find_img('nao_optante_3.png', conf=0.95):
            _escreve_relatorio_csv(f'{cnpj};Não optante;{comp};{ano}', nome=andamentos)
            print('❌ Não optante')
            p.hotkey('ctrl', 'w')
            return False
        
    _click_img('emitir_guia.png', conf=0.9)

    # Selecionar o ano e clicar em ok
    _wait_img('ano.png', conf=0.95, timeout=-1)
    time.sleep(1)

    _click_img('ano.png', conf=0.95)
    time.sleep(1)

    if _find_img('nao_optante_2.png', conf=0.95):
        _escreve_relatorio_csv(f'{cnpj};Não optante;{comp};{ano}', nome=andamentos)
        print('❌ Não optante')
        p.hotkey('ctrl', 'w')
        return False

    if not _find_img('2023.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Ano não liberado;{comp};{ano}', nome=andamentos)
        print('❌ Ano não liberado')
        p.hotkey('ctrl', 'w')
        return False

    # while not _find_img('ano_de_consulta.png', conf=0.9):
    p.press('up')
    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    _click_img('ok.png', conf=0.9)
    time.sleep(1)

    if _find_img('nao_tem_anterior.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Não consta pagamento de anos anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Não consta pagamento de anos anteriores')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('nao_tem_ano.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Não optante;{comp};{ano}', nome=andamentos)
        print('❌ Não optante')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('antes_de_prosseguir.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores;{comp};{ano}', nome=andamentos)
        print('❌ Antes de prosseguir, é necessário apresentar a(s) DASN-Simei do(s) ano(s) anteriores')
        p.hotkey('ctrl', 'w')
        return False

    # Esperar abrir a tela das guias
    _wait_img('selecione.png', conf=0.9, timeout=-1)
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
    _click_img('apurar.png', conf=0.9)
    time.sleep(2)

    if _find_img('mes_baixado.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Competência não habilitada;{comp};{ano}', nome=andamentos)
        print('❌ Competência não habilitada')
        p.hotkey('ctrl', 'w')
        return False

    if _find_img('ja_consta.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};Já consta pagamento;{comp};{ano}', nome=andamentos)
        print('❗ Já consta pagamento')
        p.hotkey('ctrl', 'w')
        return False

    _wait_img('voltar.png', conf=0.9, timeout=-1)
    return True
    

@_time_execution
def run():
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        cnpj, comp, ano, venc = empresa
        andamentos = f'Boletos MEI {comp}-{ano}'
        
        _abrir_chrome('http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao')
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
