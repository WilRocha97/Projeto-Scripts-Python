from sys import path
import datetime, time, pyautogui as p

path.append(r'..\..\_comum')
from comum_comum import _indice, _escreve_relatorio_csv, _time_execution, _open_lista_dados, _where_to_start


def esperar(imagen):
    while not p.locateOnScreen(imagen, confidence=0.9):
        time.sleep(1)


def clicar(imagen, tempo=0.1):
    try:
        p.click(p.locateCenterOnScreen(imagen, confidence=0.9))
        time.sleep(tempo)
        return True
    except:
        return False


def verificacoes(empresa):
    cod, cnpj, nome, arquivo = empresa
    if p.locateOnScreen(r'imagens\JaExiste.png'):
        p.press('enter')
        time.sleep(1)
        p.press(['esc', 'esc'])
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Arquivo já importado']), 'Importação DIRF')
        print('>>> Arquivo já importado <<<')
        return False
    if p.locateOnScreen(r'imagens\JaConsta.png'):
        p.hotkey('ctrl', 'n')
        time.sleep(1)
        p.press(['esc', 'esc'])
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Arquivo já consta na base de dados']), 'Importação DIRF')
        print('>>> Arquivo já consta na base de dados <<<')
        return False
    if p.locateOnScreen(r'imagens\Erros.png'):
        p.hotkey('alt', 'n')
        time.sleep(2)
        if p.locateOnScreen(r'imagens\Relatorio.png'):
            clicar(r'imagens\SalvarRelatorio.png')
            esperar(r'imagens\Salvar.png')
            time.sleep(0.2)
            p.write(cnpj + ' - ' + cod + '.pdf')
            time.sleep(0.2)
            p.press('enter')
            time.sleep(1)
            if p.locateOnScreen(r'imagens\Sobrescrever.png'):
                p.press('enter')
            clicar(r'imagens\FecharRelatorio.png')
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Erros ou avisos']), 'Importação DIRF')
        print('** Erros ou avisos **')
        return False
    return True


@_time_execution
def importar_dirf():
    # abrir a planilha de dados
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # configurar um indice para a planilha de dados
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    p.hotkey('win', 'm')
    time.sleep(0.5)
    p.getWindowsWithTitle('Dirf')[0].maximize()
    
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, tempos=tempos, tempo_execucao=tempo_execucao)
        
        cod, cnpj, nome, arquivo = empresa

        clicar(r'imagens\TelaInicial.png')

        # Abrir tela para selecionar o arquvio
        while p.locateOnScreen(r'imagens\TelaInicial.png'):
            p.hotkey('ctrl', 'i')
            time.sleep(3)
        while not p.locateOnScreen(r'imagens\AbrirDeclaracao.png'):
            time.sleep(1)
            p.click(p.locateCenterOnScreen(r'imagens\Lupa.png'))
        p.write(arquivo)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(1)

        if not verificacoes(empresa):
            continue

        # Esperar importar o arquivo
        esperar(r'imagens\Avancar.png')
        p.hotkey('alt', 'a')
        esperar(r'imagens\Resumo.png')

        if not verificacoes(empresa):
            continue

        if not p.locateOnScreen(r'imagens\Concluir.png'):
            ''
        p.hotkey('alt', 'n')
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Arquivo importado']), 'Importação DIRF')
        print('>>> Arquivo importado <<<')


if __name__ == '__main__':
    try:
        importar_dirf()
    except:
        p.alert(title='Script incrível', text='ERRO')
