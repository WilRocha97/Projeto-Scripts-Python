# -*- coding: utf-8 -*-
import os, time, pyperclip, shutil, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _time_execution, _escreve_relatorio_csv, ask_for_dir


def abre_declaracao(documentos, arquivo):
    # aguarda a tela de importar declaração
    _wait_img('importar.png', conf=0.9)
    time.sleep(1)
    _click_img('importar.png', conf=0.9)
    
    # aguarda a tela de selecionar arquivo para importar
    while not _find_img('importacao.png', conf=0.9):
        if _find_img('importar_declaracao.png', conf=0.8):
            _click_img('importar_declaracao.png', conf=0.8)
        if _find_img('importar_declaracao_anual.png', conf=0.8):
            _click_img('importar_declaracao_anual.png', conf=0.8)
        if _find_img('importar_declaracao_anual_2.png', conf=0.8):
            _click_img('importar_declaracao_anual_2.png', conf=0.8)
        if _find_img('importar_declaracao.png', conf=0.8):
            break
            
    time.sleep(2)
    # insere o caminho do arquivo que será importado
    caminho_arquivo = str(os.path.join(documentos, arquivo))
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
            return False , 'Erro na importação da Declaração. Este arquivo está corrompido', caminho_arquivo
    
    # cofirma que foi importado
    p.press('enter')
    return True, 'ok', caminho_arquivo


def abre_recibo():
    # aguarda o botão de abrir declaração
    _wait_img('abrir_recibo.png', conf=0.9)
    time.sleep(1)
    _click_img('abrir_recibo.png', conf=0.9)
    
    # aguarda a tela de imprimir declaração
    while not _find_img('impressao_recibo.png', conf=0.9):
        if _find_img('impressao_recibo_2.png', conf=0.9):
            break
    time.sleep(1)
    
    # clica no arquivo
    _click_position_img('seleciona_arquivo.png', '+', pixels_y=28, conf=0.9)
    
    time.sleep(1)
    # clica em ok
    p.hotkey('alt', 'o')


def imprimir_arquivo():
    # aguarda o botão de imprimir delcaração
    while not _find_img('imprimir.png', conf=0.9):
        _click_img('abrir_declaracao.png', conf=0.9, timeout=1)
        time.sleep(1)
        _click_img('abrir_declaracao_2.png', conf=0.9, timeout=1)
        time.sleep(1)
        if _find_img('imprimir_2.png', conf=0.9):
            break
        
    time.sleep(2)
    p.hotkey('ctrl', 'p')

    # aguarda a tela de impressão
    while not _find_img('selecao_de_impressao.png', conf=0.9):
        if _find_img('selecao_de_impressao_2.png', conf=0.9):
            break
    time.sleep(1)
    
    _click_img('toda_a_declaracao.png', conf=0.9)
    time.sleep(1)
    
    p.hotkey('alt', 'o')
    
    
def salvar_pdf():
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
        if _find_img('sobrescrever.png', conf=0.9):
            p.hotkey('alt', 'o')
        if _find_img('sobrescrever_2.png', conf=0.9):
            p.hotkey('alt', 'o')

    time.sleep(3)
    if _find_img('fechar.png', conf=0.9):
        _click_position_img('fechar.png', '+', pixels_x=911, conf=0.9)
    elif _find_img('fechar_2.png', conf=0.9):
        _click_position_img('fechar_2.png', '+', pixels_x=911, conf=0.9)
    
    return True


def mover_pdf(final_folder_pdf):
    download_folder = 'C:\\Users\\robo\\Documents'
    
    for arquivo in os.listdir(download_folder):
        if arquivo.endswith('.pdf'):
            shutil.move(os.path.join(download_folder, arquivo), os.path.join(final_folder_pdf, arquivo))


def exclui_declaracao(andamentos):
    _click_img('excluir.png', conf=0.9)
    while not _find_img('excluir_declaracao.png', conf=0.9):
        time.sleep(1)
        if _find_img('divergencia_1.png', conf=0.9):
            _escreve_relatorio_csv(f'A soma de todos os rendimentos é menor que a das deduções', nome=andamentos, end=';')
            p.hotkey('alt', 's')
            
    time.sleep(1)
    p.hotkey('alt', 't')
    time.sleep(1)
    p.hotkey('alt', 'o')
    time.sleep(1)
    p.hotkey('alt', 's')
    time.sleep(1)
    

@_time_execution
def run():
    tipo_arquivo = p.confirm(buttons=['Declarações', 'Recibos'])
    documentos = ask_for_dir()
    # pega o nome para a pasta final do aqruivo
    pasta = documentos.split('/')
    # define as pastas do programa do irpf
    irpf_folder = f'C:\\Arquivos de Programas RFB\\IRPF{pasta[5]}\\transmitidas'
    
    if tipo_arquivo == 'Declarações':
        pasta = documentos.split('/')
        andamentos = f'Declarações IRPF {pasta[6]}'
    
        final_folder_pdf = f'C:\\Users\\robo\\Documents\\Declarações\\{pasta[6]}_PDF'
        final_folder = f'C:\\Users\\robo\\Documents\\Declarações\\{pasta[6]}'
    
        os.makedirs(final_folder_pdf, exist_ok=True)
        os.makedirs(final_folder, exist_ok=True)
    
    else:
        # cria o nome da planilha de andamentos
        andamentos = f'Recibos IRPF {pasta[6]}'
        
        # define a pasta final do PDF e a pasta final dos arquivos
        final_folder_pdf = f'C:\\Users\\robo\\Documents\\Recibos\\{pasta[6]}_PDF'
        final_folder = f'C:\\Users\\robo\\Documents\\Recibos\\{pasta[6]}'
    
        os.makedirs(final_folder_pdf, exist_ok=True)
        os.makedirs(final_folder, exist_ok=True)
        
        if os.listdir(irpf_folder) != '[]':
            for arquivo in os.listdir(irpf_folder):
                shutil.move(os.path.join(irpf_folder, arquivo), os.path.join(documentos, arquivo))
        
    # for cnpj in cnpjs:
    for arquivo in os.listdir(documentos):
        if tipo_arquivo == 'Declarações':
            
            if arquivo.endswith('.DEC'):
                print(f'\n{arquivo}')
                funcao, situacao, caminho = abre_declaracao(documentos, arquivo)
            else:
                shutil.move(os.path.join(documentos, arquivo), os.path.join(final_folder, arquivo))
                continue
        
            if funcao:
                imprimir_arquivo()
                salvar_pdf()
                mover_pdf(final_folder_pdf)
                
                print('✔ PDF gerado com sucesso')
                shutil.move(os.path.join(documentos, arquivo), os.path.join(final_folder, arquivo))
                _escreve_relatorio_csv(f'{arquivo};PDF gerado com sucesso', nome=andamentos, end=';')
                
                exclui_declaracao(andamentos)
                
            else:
                print(f'❌ {situacao}')
                shutil.move(os.path.join(documentos, arquivo), os.path.join(final_folder, arquivo))
                _escreve_relatorio_csv(f'{arquivo};{situacao}', nome=andamentos, end=';')
        
        else:
            if not arquivo.endswith('.REC'):
                continue

            # define o nome do arquivo do recibo e do arquivo da declaração
            arquivo_recibo = arquivo
            arquivo_declaracao = arquivo.replace('.REC', '.DEC')

            # move os arquivos da declaração e do recibo para a pasta de transmitidas do programa do irpf para poder gerar o PDF do recibo
            shutil.move(os.path.join(documentos, arquivo_declaracao), os.path.join(irpf_folder, arquivo_declaracao))
            shutil.move(os.path.join(documentos, arquivo_recibo), os.path.join(irpf_folder, arquivo_recibo))

            print(f'\n{arquivo}')
            abre_recibo()
            salvar_pdf()
            mover_pdf(final_folder_pdf)
            
            # move os arquivos da declaração e do recibo para a pasta final
            shutil.move(os.path.join(irpf_folder, arquivo_declaracao), os.path.join(final_folder, arquivo_declaracao))
            shutil.move(os.path.join(irpf_folder, arquivo_recibo), os.path.join(final_folder, arquivo_recibo))

            print('✔ PDF gerado com sucesso')
            _escreve_relatorio_csv(f'{arquivo};PDF gerado com sucesso', nome=andamentos, end=';')

        _escreve_relatorio_csv('', nome=andamentos)
            
if __name__ == '__main__':
    run()
