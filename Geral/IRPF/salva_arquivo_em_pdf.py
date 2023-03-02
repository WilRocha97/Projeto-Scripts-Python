# -*- coding: utf-8 -*-
import os, time, pyperclip, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def abre_arquivo(arquivo):
    # aguarda o botão de nova delcaração
    if not _find_img('importar.png', conf=0.9):
        _wait_img('nova.png', conf=0.9)
        time.sleep(1)
        _click_img('nova.png', conf=0.9)
    
    # aguarda a tela de importar declaração
    _wait_img('importar.png', conf=0.9)
    time.sleep(1)
    _click_img('importar.png', conf=0.9)
    
    # aguarda a tela de selecionar arquivo para importar
    _wait_img('importacao.png', conf=0.9)
    time.sleep(2)
    
    # insere o caminho do arquivo que será importado
    caminho_arquivo = 'V:\Setor Robô\Scripts Python\Geral\IRPF\ignore\Arquivos\InsiraAqui'.replace('InsiraAqui', str(arquivo))
    pyperclip.copy(caminho_arquivo)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    
    p.press('enter')
    time.sleep(1)
    
    # enquanto o arquivo não importar verifica se já existe o mesmo arquivo importado, se sim, confirma a reimportação
    while not _find_img('importou_sucesso.png', conf=0.9):
        if _find_img('ja_existe_declaracao.png', conf=0.9):
            p.hotkey('alt', 's')
        
        if _find_img('corrompido.png', conf=0.9):
            p.press('enter')
            return False , 'Erro na importação da Declaração. Este arquivo está corrompido.'
    
    # cofirma que foi importado
    p.press('enter')
    return True, 'ok'


def salva_pdf():
    # aguarda o botão de imprimir delcaração
    _wait_img('imprimir.png', conf=0.9)
    time.sleep(1)
    _click_img('imprimir.png', conf=0.9)

    # aguarda a tela de impressão
    _wait_img('selecao_de_impressao.png', conf=0.9)
    time.sleep(1)
    p.hotkey('alt', 'o')

    # aguarda a tela de visualização
    _wait_img('salvar_pdf.png', conf=0.9)
    time.sleep(1)
    _click_img('salvar_pdf.png', conf=0.9)

    # aguarda a tela de salvar
    _wait_img('tela_salvar.png', conf=0.9)
    time.sleep(1)
    p.press('enter')
    
    while _find_img('tela_salvar.png', conf=0.9):
        time.sleep(1)
        
    time.sleep(1)
    if _find_img('sobrescrever.png', conf=0.9):
        p.hotkey('alt', 'o')

    time.sleep(3)
    _click_position_img('fechar.png', 'soma', 911, conf=0.9)
    
    return True


def mover_pdf():
    download_folder = 'C:\\Users\\robo\\Documents'
    final_folder = "V:\\Setor Robô\\Scripts Python\\Geral\\IRPF\\execução\\Declarações"
    
    for arquivo in os.listdir(download_folder):
        if arquivo.endswith('.pdf'):
            shutil.move(os.path.join(download_folder, arquivo), os.path.join(final_folder, arquivo))


@_time_execution
def run():
    os.makedirs('execução/Declarações', exist_ok=True)
    empresas = _open_lista_dados()
    andamentos = 'Declarações IRPF'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        
        indice, arquivo = empresa
        
        funcao, situacao = abre_arquivo(arquivo)
        
        if not funcao:
            print(f'❌ {situacao}')
            _escreve_relatorio_csv(f'{indice};{arquivo};{situacao}', nome=andamentos)
            continue
            
        salva_pdf()
        mover_pdf()
        
        print('✔ PDF gerado com sucesso')
        _escreve_relatorio_csv(f'{indice};{arquivo};PDF gerado com sucesso', nome=andamentos)

        
    
    
if __name__ == '__main__':
    run()
