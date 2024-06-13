import datetime, fitz, time, os, pyperclip, pyautogui as p

from sys import path
path.append(r'..\..\_comum')
from meu_inss_comum import _login_gov
from chrome_comum import _abrir_chrome, _acessar_site_chrome
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def busca_analise():
    _acessar_site_chrome('https://meu.inss.gov.br/#/extrato-previdenciario')
    print('>>> Aguardando extratos')
    # aguarda a tela da simulação carregar
    while not _find_img('tela_extratos.png', conf=0.9):
        if _find_img('erro_buscar_dados.png', conf=0.9):
            p.hotkey('ctrl', 'w')
            return False, 'recomeçar'
        time.sleep(1)

    return True, ''


def salva_arquivo(caminho, arquivo_nome, arquivo):
    print(f'>>> Salvando {arquivo.lower()}')
    
    # aguarda a tela de salvar do navegador abrir
    while not _find_img('salvar_como.png', conf=0.9):
        if _find_img('erro_salvar_extrato.png', conf=0.9):
            return False
        time.sleep(1)
            
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    pyperclip.copy(arquivo_nome)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy(caminho)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    if _find_img('substituir.png', conf=0.95):
        p.press('s')
    time.sleep(5)
    
    while True:
        try:
            doc = fitz.open(os.path.join(caminho, arquivo_nome + '.pdf'), filetype="pdf")
            break
        except:
            pass
    
    print(f'✔ {arquivo} baixado com sucesso')
    time.sleep(1)
    
    return True


@_time_execution
@_barra_de_status
def run(window):
    tempos = [datetime.datetime.now()]
    tempo_execucao = []
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        tempos, tempo_execucao = _indice(count, total_empresas, empresa, index, window, tempos, tempo_execucao)

        cpf, nome, senha = empresa
        nome = nome.strip()
        
        andamentos = f'Extratos de Contribuições - Previdenciárias e Remunerações'

        _abrir_chrome('https://meu.inss.gov.br/#/login', iniciar_com_icone=True)
        # faz login no cpf
        while True:
            resultado = _login_gov(cpf, senha)
            if resultado != 'recomeçar':
                break

        if resultado == 'ok':
            # busca o extrato
            while True:
                situacao, resultado = busca_analise()
                if resultado == 'recomeçar':
                    while True:
                        resultado = _login_gov(cpf, senha)
                        if resultado != 'recomeçar':
                            break
                else:
                    break

            if situacao:
                if not _find_img('baixar_pdf.png', conf=0.95):
                    p.press('end')

                time.sleep(1)
                _click_img('baixar_pdf.png', conf=0.95)

                while not _find_img('previdenciarias_remuneracoes.png', conf=0.9):
                    time.sleep(1)

                _click_img('previdenciarias_remuneracoes.png', conf=0.95)

                caminho = 'V:\Setor Robô\Scripts Python\Meu INSS\Extrato de Contribuição CNIS\execução\Extratos'
                os.makedirs('execução\Extratos', exist_ok=True)
                arquivo_nome = f'CNIS Previdenciárias e Remunerações - {cpf} - {nome}'
                if salva_arquivo(caminho, arquivo_nome, 'Extrato CNIS com relações previdenciárias e remunerações'):
                    resultado = 'Extrato salvo'
                else:
                    resultado = 'Erro no site, Não salvou extrato'

        _escreve_relatorio_csv(f'{cpf};{nome};{senha};{resultado}', nome=andamentos)
        p.hotkey('ctrl', 'w')


if __name__ == '__main__':
    empresas = _open_lista_dados()
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
