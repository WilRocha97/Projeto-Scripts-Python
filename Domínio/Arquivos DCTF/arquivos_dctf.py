# -*- coding: utf-8 -*-
import os
import time, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status
from dominio_comum import _login_web, _abrir_modulo, _login


def arquivos_dctf(empresa, periodo, andamento):
    cod, cnpj, nome, regime, movimento = empresa
    nome_arquivo = 'M:\DCTF_{}.RFB'.format(cod)
    
    # aguarda a tela do domínio
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # relatórios
    p.hotkey('alt', 'r')
    
    time.sleep(0.5)
    # informativos
    p.press('n')
    time.sleep(0.5)
    # federais
    p.press('f')
    time.sleep(0.5)
    # dctf
    p.press('d')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    # mensal
    p.press('m')
    
    # aguarda a tela abrir
    while not _find_img('dctf_mensal.png', conf=0.9):
        if _find_img('dctf_mensal_2.png', conf=0.9):
            break
        time.sleep(1)

    # digita o período
    p.write(periodo)
    time.sleep(1)
    
    # desce para inserir o nome do arquivo
    p.press('tab')
    time.sleep(1)
    
    # apaga o nome que estiver digitado
    p.press('delete', presses=50)
    time.sleep(1)
    
    # digita o nome do arquivo
    p.write(nome_arquivo)
    time.sleep(1)
    
    # abre outros dados
    p.hotkey('alt', 'o')
    
    # aguarda a tela abrir
    while not _find_img('outros_dados.png', conf=0.9):
        if _find_img('outros_dados_2.png', conf=0.9):
            break
        
        # verifica a competência
        if _find_img('data_inicio.png', conf=0.9):
            _click_img('data_inicio.png', conf=0.9)
            p.press('enter')
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, 'Competência informada menor que a data de início efetivo das atividades.']), nome=andamento)
            print('❗ Competência informada menor que a data de início efetivo das atividades.')
            time.sleep(1)
            p.press('esc', presses=5)
            return 'ok', nome_arquivo
        
        time.sleep(2)
    
    if p.locateOnScreen(r'imgs/declarante_nao_contribui_irpj.png'):
        regime = 'Declarante não é contribuinte do IRPJ'
    else:
        # configura o regime, forma de tributação
        if regime == 'Lucro Real':
            if not p.locateOnScreen(r'imgs/real_estimativa.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_real_estimativa.png', conf=0.99)
                p.press('enter')

        if regime == 'Lucro Presumido':
            if not p.locateOnScreen(r'imgs/presumido.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_presumido.png', conf=0.99)
                p.press('enter')
                
        if regime == 'Imunes':
            if not p.locateOnScreen(r'imgs/imune.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_imune.png', conf=0.99)
                p.press('enter')
            
        if regime == 'Isentas':
            if not p.locateOnScreen(r'imgs/isenta.png'):
                p.click(1215, 192)
                time.sleep(0.5)
                _click_img('seleciona_isenta.png', conf=0.99)
                p.press('enter')
            
        if regime == 'Simples Nacional':
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, 'Empresa Simples Nacional não gera arquivo.']), nome=andamento)
            print('❗ Empresa Simples Nacional não gera arquivo.')
            return 'ok', nome_arquivo

        # configura o regime, qualificação
        time.sleep(1)
        p.click(1215, 217)
        time.sleep(0.5)
        # pj em geral
        p.write('PJ em Geral')
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)
    
    # levantou balancete/balanço de suspensão e redução
    if regime == 'Lucro Real':
        if not p.locateOnScreen(r'imgs/levantou_marcado.png'):
            p.click(414, 421)
            time.sleep(0.5)
    else:
        if p.locateOnScreen(r'imgs/levantou_marcado.png'):
            p.click(414, 421)
            time.sleep(0.5)
        
    # situação da pj no mês da declaração
    p.click(1214, 488)
    time.sleep(0.5)
    # pj não se enquadra em nenhuma das situações anteriores no mês da declaração
    p.write('PJ nao')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(1)
    
    # criterio de reconhecimento das variações monetárias dos direitos de crédito e das obrigações do contribuinte, em função da taxa de câmbio
    if not p.locateOnScreen(r'imgs/criterio_marcado.png'):
        p.click(413, 541)
        time.sleep(1)
    
    if movimento == 'Sem movimento':
        if not p.locateOnScreen(r'imgs/pj_inativa_marcado.png'):
            p.click(412, 601)
            time.sleep(1)
    
    # sem alteração do regime
    p.click(1215, 547)
    time.sleep(0.5)
    if _find_img('sem_alteracao.png', conf=0.99):
        _click_img('sem_alteracao.png', conf=0.99)
        time.sleep(1)
    
    # ok
    p.hotkey('alt', 'o')
    time.sleep(1)
    
    # aguarda a tela fechar
    while _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)
    
    # exportar arquivo
    p.hotkey('alt', 'x')
    
    print('>>> Gerando arquivo')
    # aguarda o arquivo gerar
    while _find_img('dctf_mensal.png', conf=0.9) or _find_img('dctf_mensal_2.png', conf=0.9):
        time.sleep(1)
        ocorrencias = [('nao_gerou_arquivo.png', 'Não gerou arquivo', 'ok'),
                       ('saldo_nao_calculado.png', 'Saldo dos impostos não foi calculado no período;' + periodo, 'ok'),
                       ('nao_tem_parametro.png', 'Não existe parametro para a vigência;' + periodo, 'ok'),
                       ('exportacao_cancelada.png', 'Exportação cancelada', 'ok'),
                       ('final_da_exportacao.png', 'Arquivo gerado', 'arquivo gerado')]
        
        for ocorrencia in ocorrencias:
            if _find_img(ocorrencia[0], conf=0.9):
                p.press('enter')
                _escreve_relatorio_csv(';'.join([cod, cnpj, nome, regime, ocorrencia[1]]), nome=andamento)
                print(f'❗ {ocorrencia[1]}')
                time.sleep(1)
                p.press('esc', presses=5)
                return ocorrencia[2], nome_arquivo
        
        if _find_img('imune_irpj.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('enter')
            mudar_regime('Imunes')
            regime = f'AE: {regime}, Domínio: Imunes'

        if _find_img('isenta_alerta.png', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('enter')
            mudar_regime('Isentas')
            regime = f'AE: {regime}, Domínio: Isentas'
        
    print('❗ Erro inesperado')
    p.press('esc', presses=5)
    time.sleep(3)
    return 'ok', nome_arquivo


def mudar_regime(regime_certo):
    # abre outros dados
    p.hotkey('alt', 'o')
    
    # aguarda a tela abrir
    while not _find_img('outros_dados.png', conf=0.9):
        if _find_img('outros_dados_2.png', conf=0.9):
            break
    
    if regime_certo == 'Imunes':
        if not p.locateOnScreen(r'imgs/imune.png'):
            p.click(1215, 192)
            time.sleep(0.5)
            _click_img('seleciona_imune.png', conf=0.99)
            p.press('enter')
    
    if regime_certo == 'Isentas':
        if not p.locateOnScreen(r'imgs/isenta.png'):
            p.click(1215, 192)
            time.sleep(0.5)
            _click_img('seleciona_isenta.png', conf=0.99)
            p.press('enter')
        
    # ok
    p.hotkey('alt', 'o')
    time.sleep(1)
    
    # aguarda a tela fechar
    while _find_img('outros_dados.png', conf=0.9):
        time.sleep(1)
    
    # exportar arquivo
    p.hotkey('alt', 'x')
    
    
def mover_arquivo(nome_arquivo):
    nome_arquivo = nome_arquivo.replace('M:\DCTF', 'DCTF')
    os.makedirs('execução/Arquivos', exist_ok=True)
    final_folder = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Arquivos DCTF\\execução\\Arquivos'
    folder = 'C:\\'
    shutil.move(os.path.join(folder, nome_arquivo), os.path.join(final_folder, nome_arquivo))


@_time_execution
@_barra_de_status
def run(window):
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)
        
        while True:
            if not _login(empresa, andamentos):
                break
            else:
                resultado, nome_arquivo = arquivos_dctf(empresa, periodo, andamentos)
                if resultado == 'arquivo gerado':
                    mover_arquivo(nome_arquivo)
                    break
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
    
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
                
                if resultado == 'ok':
                    break
    

if __name__ == '__main__':
    p.mouseInfo()
    periodo = p.prompt(text='Qual o período do arquivo', title='Script incrível', default='00/0000')
    empresas = _open_lista_dados()
    andamentos = 'Arquivos para DARF DCTF'
    
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        