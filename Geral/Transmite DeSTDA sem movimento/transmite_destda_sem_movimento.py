# -*- coding: utf-8 -*-
import pyperclip, time, os, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados SEDIF.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')


def reinstala_sedif():
    time.sleep(2)
    p.hotkey('win', 'm')
    os.startfile(r'C:\Users\robo\Downloads\sedif_Setup')
    _wait_img('instala_sedif.png', conf=0.9, timeout=-1)
    p.press('enter', presses=5, interval=1)
    _wait_img('final_instala_sedif.png', conf=0.9, timeout=-1)
    p.press('enter')
    _wait_img('atualiza_sedif.png', conf=0.9, timeout=-1)
    _click_img('atualiza_sedif.png', conf=0.9)
    time.sleep(1)
    _click_img('nao.png', conf=0.9)
    time.sleep(5)
    '''if _find_img('ja_atualizado.png', conf=0.9):
        p.press('enter')
        time.sleep(1)
        p.press('esc')'''
    time.sleep(1)
    

def configura(empresa, comp):
    cnpj, nome = empresa
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]
    p.moveTo(100, 100)
    
    _wait_img('novo_documento.png', conf=0.9, timeout=-1)
    _click_img('novo_documento.png', conf=0.9)
    
    while not _find_img('nome_empresa.png', conf=0.9):
        time.sleep(10)
        p.click(900, 900)
    _click_img('nome_empresa.png', conf=0.9)
    
    p.write(nome)
    time.sleep(5)
    
    while True:
        try:
            p.hotkey('ctrl', 'a')
            time.sleep(0.5)
            p.hotkey('ctrl', 'c')
            time.sleep(0.5)
            text = pyperclip.paste()
            break
        except:
            pass
    print(text)
    
    if text != nome:
        _escreve_relatorio_csv(f'{cnpj};{nome};{text};Empresa não encontrada no SEDIF', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
        print('❗ Empresa não encontrada')
        p.press('esc')
        p.hotkey('alt', 'l')
        _wait_img('novo.png', conf=0.9, timeout=-1)
        p.hotkey('alt', 'f')
        time.sleep(1)
        p.hotkey('alt', 'f4')
        time.sleep(1)
        p.hotkey('alt', 's')
        return False
    
    p.press('enter')
    p.write(comp)
    p.hotkey('alt', 'p')
    
    while not _find_img('escrituracao.png', conf=0.99):
        time.sleep(1)
    p.write('DeSTDA')
    p.press('tab', presses=2)
    p.write('Original')
    p.press('tab')
    p.write('Sem dados informados')
    time.sleep(0.5)
    _click_img('confirmar.png', conf=0.9)
    time.sleep(20)
    
    if _find_img('ja_existe_declaracao.png', conf=0.9):
        p.press('esc')
    
    if _find_img('preenchimento_errado.png', conf=0.9):
        _escreve_relatorio_csv(f'{cnpj};{nome};Preenchimento incorreto dos campos {comp}', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
        print('❗ Preenchimento incorreto dos campos')
        p.press('enter')
        p.hotkey('alt', 'l')
        time.sleep(2)
        p.hotkey('alt', 'f')
        return False
    
    p.hotkey('alt', 'f')
    return True


def excluir_documento():
    while not _find_img('abrir_arquivo.png', conf=0.9):
        time.sleep(1)
        p.press('up')
        _click_img('minimizar.png', conf=0.9)
    
    time.sleep(1)
    p.hotkey('alt', 'e')
    time.sleep(1)
    p.hotkey('alt', 's')
    
    # espera excluir o arquivo
    _wait_img('arquivo_excluido.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'f')
    

def transmitir(empresa, comp):
    # p.alert(text='alerta')
    cnpj, nome = empresa
    mes = comp.split('/')[0]
    ano = comp.split('/')[1]

    p.hotkey('alt', 'f')
    
    # abrir arquivo da empresa
    while not _find_img('abrir_arquivo.png', conf=0.9):
        _wait_img('abrir_documento.png', conf=0.9, timeout=-1)
        if _find_img('abrir_documento.png', conf=0.9):
            _click_img('abrir_documento.png', conf=0.9)
        time.sleep(3)
        p.press('up')
        _click_img('minimizar.png', conf=0.9)
    p.press('enter')
    time.sleep(2)
    p.hotkey('alt', 'b')
    
    # espera tela inicial do sedif
    _wait_img('tela_inicial.png', conf=0.9, timeout=-1)
    
    # clicar em na aba encerrar
    _click_img('encerrar.png', conf=0.9)
    time.sleep(2)
    
    # gerar documento
    p.hotkey('alt', 'g')
    
    # esperar o botão para gerar assinatura do documento
   
    _wait_img('assinatura.png', conf=0.9)
    time.sleep(1)
    
    p.hotkey('alt', 'i')

    # espera o botão de transmitir
    while not _find_img('transmitir.png', conf=0.9):
        if _find_img('criticas.png', conf=0.9):
            p.press('enter', presses=4, interval=2)
            time.sleep(1)
            p.hotkey('alt', 'f')
            time.sleep(1)
            _click_img('iniciar.png', conf=0.9)
            p.hotkey('alt', 'f')
            time.sleep(1)
            p.press('f')
            _escreve_relatorio_csv(f'{cnpj};{nome};Existe criticas nos dados que impedem a assinatura', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
            print('❌ Existe criticas nos dados que impedem a assinatura')
            
            excluir_documento()
            
            return False
            
        if _find_img('assinar_com_certificado.png', conf=0.9):
            p.hotkey('alt', 'n')
        if _find_img('processo_finalizado_ass.png', conf=0.9):
            p.press('enter')
        if _find_img('erro_create_file.png', conf=0.9):
            p.press('enter')
        time.sleep(1)
    
    p.hotkey('alt', 't')
    
    # iniciar transmissão do arquivo
    _wait_img('transmissao.png', conf=0.9, timeout=-1)
    time.sleep(1)
    p.hotkey('alt', 'i')
    
    while not _find_img('identificacao.png', conf=0.9):
        time.sleep(1)
        if _find_img('access_violation.png', conf=0.9):
            p.press('enter')
            
    _click_img('identificacao.png', conf=0.9)

    # usuário e senha do contribuinte para transmitir sedif sem movimento
    p.write(user[0])
    p.press('tab')
    p.write(user[1])
    _click_img('ok.png', conf=0.9)
    
    excecoes = [('cpf_do_responsavel_invalido.png', 'CPF do responsável inválido'),
                 ('cnpj_informado_nao_confere.png', 'CNPJ informado não confere com o existente'),
                 ('usuario_nao_cadastrado.png', 'Usuário não cadastrado'),
                 ('arquivo_corrompido.png', 'Arquivo de configuração corrompido'),
                 ('erro.png', 'Erro no SEDIF')]
    
    while not _find_img('processo_finalizado.png', conf=0.9):
        time.sleep(1)
        for excecao in excecoes:
            
            if _find_img(excecao[0], conf=0.9):
                if not excecao[1] == 'Erro no SEDIF':
                    _escreve_relatorio_csv(f'{cnpj};{nome};{excecao[1]}', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
                else:
                    p.press('enter')
                    
                print(f'❌ {excecao[1]}')
                
                p.hotkey('alt', 'f')
                time.sleep(1)
                p.hotkey('alt', 'f')
                time.sleep(1)
                p.press('f')
                
                excluir_documento()
                return excecao[1]

    p.press('enter')
    
    _wait_img('salvar_recibo.png', conf=0.9, timeout=-1)
    time.sleep(1)

    _click_img('salvar_recibo.png', conf=0.9)

    _wait_img('salvar_como.png', conf=0.9, timeout=-1)
    time.sleep(2)
    
    _click_img('insere_nome.png', conf=0.9)
    _click_img('insere_nome_2.png', conf=0.9, timeout=2)
    
    while True:
        try:
            pyperclip.copy(f'Recibo de entrega SN - {mes}-{ano} - {cnpj}')
            time.sleep(1)
            p.hotkey('ctrl', 'v')
            break
        except:
            pass
    time.sleep(2)
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    
    os.makedirs(r'V:\Setor Robô\Scripts Python\Geral\Transmite DeSTDA sem movimento\execução\Recibos', exist_ok=True)
    pyperclip.copy('V:\Setor Robô\Scripts Python\Geral\Transmite DeSTDA sem movimento\execução\Recibos')
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    
    p.press('enter')
    time.sleep(1)
    p.hotkey('alt', 'l')
    time.sleep(1)
    
    _wait_img('deseja_visualizar.png', conf=0.9, timeout=-1)
    _click_img('deseja_visualizar.png', conf=0.9)
    time.sleep(1)
    
    p.press('right')
    p.press('enter')

    _wait_img('recibo_gerado.png', conf=0.9, timeout=-1)
    time.sleep(1)

    p.press('enter')

    _wait_img('transmissao.png', conf=0.9, timeout=-1)
    _click_img('transmissao.png', conf=0.9)
    time.sleep(2)

    p.hotkey('alt', 'f')
    time.sleep(1)
    p.hotkey('alt', 'f')
    time.sleep(1)
    p.hotkey('alt', 'f')
    time.sleep(1)

    _wait_img('tela_inicial.png', conf=0.9, timeout=-1)

    _wait_img('iniciar.png', conf=0.9, timeout=-1)
    _click_img('iniciar.png', conf=0.9)
    p.hotkey('alt', 'f')
    time.sleep(1)
    p.press('f')
    
    excluir_documento()
    
    _escreve_relatorio_csv(f'{cnpj};{nome};Competência transmitida', nome=f'Transmite DeSTDA sem movimento {mes} - {ano}')
    print('✔ Competência transmitida')
    
    return True


@_time_execution
@_barra_de_status
def run(window):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count + index)} de {str(len(total_empresas) + index)} | {str((len(total_empresas) + index) - (count + index))} Restantes')
        _indice(count, total_empresas, empresa, index)
        
        erro = 'sim'
        while erro == 'sim':
            reinstala_sedif()
            if not configura(empresa, comp):
                break
            erro = transmitir(empresa, comp)
            p.hotkey('alt', 'f4')
            time.sleep(1)
            p.hotkey('alt', 's')
            
            if erro == 'Erro no SEDIF':
                erro = 'sim'
            else:
                erro = 'não'


if __name__ == '__main__':
    comp = p.prompt(title='Script incrível', text='Qual competência', default='00/0000')
    empresas = _open_lista_dados()
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
