# -*- coding: utf-8 -*-
import fitz, re, time, os, pyautogui as p, datetime as date_time
from datetime import datetime

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start
from dominio_comum import _login, _login_web, _abrir_modulo, _salvar_pdf


def escreve_dados(cod, nome):
    # cria a planilha de dados
    f = open(os.path.join('ignore', 'Dados.csv'), 'a', encoding='latin-1')
    f.write(f'{cod};CNPJ;{nome}\n')
    f.close()
    
    
def open_lista_dados(file, encode='latin-1'):
    # abre a planilha de dados
    try:
        with open(file, 'r', encoding=encode) as f:
            dados = f.readlines()
    except Exception as e:
        p.alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    print('>>> usando dados de ' + file.split('/')[-1])
    return list(map(lambda x: tuple(x.replace('\n', '').split(';')), dados))


def gera_arquivo(comp, andamento, cod='*', cnpj='', nome=''):
    # espera o botão de relatórios do domínio aparecer na tela
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    
    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('gerenciador_de_relatorios.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'r')
        time.sleep(0.5)
        # gerador de relatórios
        p.press('i', presses=2)
        time.sleep(0.5)
        p.press('enter')
        time.sleep(2)
    
    # escreve o nome da opção para selecionar
    time.sleep(0.5)
    p.write('Relação de Empregados')
    time.sleep(0.5)
    p.press('enter')
    # espera aparecer o tipo do relatório que sera usado e depois clica nele
    _wait_img('relatorio_modelo_veiga.png', conf=0.9, timeout=-1)
    _click_img('relatorio_modelo_veiga.png', conf=0.9)
    
    # insere o código da empresa, '*' para selecionar todas
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(cod)
    
    # insere '*' para selecionar todos os funcionários
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.press('*')

    hoje, data_subtraida = define_data()

    # insere a data de início e fim do contrato
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(data_subtraida)

    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('del', presses=4)
    time.sleep(0.5)
    p.write(hoje)
    
    # executa
    time.sleep(0.5)
    p.hotkey('alt', 'e')
    
    # enquanto o relatório não é gerado, verifica se aparece a mensagem dizendo que não possuí dados para emitir
    while not _find_img('contrato_experiencia.png', conf=0.9):
        time.sleep(1)
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, nome, 'Sem dados para emitir']), nome=andamento)
            print('❌ Sem dados para emitir')
            p.press('enter')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False
    
    # se gerar o relatório para todas as empresas só salva o pdf, se não tenta enviar o arquivo para o cliente
    if cod == '*':
        _salvar_pdf()
    else:
        if not _find_img('enviar_arquivo.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, nome, 'Não possuí opção de enviar o relatório para o cliente']), nome=andamento)
            print('❌ Não possuí opção de enviar o relatório para o cliente')
            p.press('esc', presses=5)
            return False
            
        envia_experiencia(comp)
        _escreve_relatorio_csv(';'.join([cod, nome, 'Relatório enviado']), nome=andamento)
        print('✔ Relatório enviado')

    # fechar qualquer possível tela aberta
    p.press('esc', presses=5)
    time.sleep(2)
    return True
    

def pega_empresas_com_exp():
    # Definir os padrões de regex
    padraozinho_nome1 = re.compile(r'Local\n(\d) - (.+)\n')
    padraozinho_nome2 = re.compile(r'Local\n(\d\d) - (.+)\n')
    padraozinho_nome3 = re.compile(r'Local\n(\d\d\d) - (.+)\n')
    padraozinho_nome4 = re.compile(r'Local\n(\d\d\d\d) - (.+)\n')
    prevtexto_nome = ''
    
    # para cada página do pdf
    if os.path.exists(os.path.join('ignore', 'Dados.csv')):
        os.remove(os.path.join('ignore', 'Dados.csv'))
        
    # abre o pdf gerado no domínio com todas as empresas que possuem experiência a vencer
    with fitz.open(os.path.join('ignore', 'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga.pdf')) as pdf:
        for page in pdf:
            andamento = f'Pagina = {str(page.number + 1)}'
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # Procura o nome da empresa no texto do pdf
                matchzinho_nome = padraozinho_nome1.search(textinho)
                if not matchzinho_nome:
                    matchzinho_nome = padraozinho_nome2.search(textinho)
                    if not matchzinho_nome:
                        matchzinho_nome = padraozinho_nome3.search(textinho)
                        if not matchzinho_nome:
                            matchzinho_nome = padraozinho_nome4.search(textinho)
                            if not matchzinho_nome:
                                continue
    
                # Guardar o nome da empresa
                matchtexto_nome = matchzinho_nome.group(2)
                # Guardar o código da empresa no DPCUCA
                matchtexto_cod = matchzinho_nome.group(1)
    
                if matchtexto_nome == prevtexto_nome:
                    continue
                    
                escreve_dados(matchtexto_cod, matchtexto_nome)
                prevtexto_nome = matchtexto_nome
            except:
                _escreve_relatorio_csv(andamento, nome='Erros')
                continue
    
    # seleciona a planilha gerada para usar no script
    empresas = open_lista_dados(os.path.join('ignore', 'Dados.csv'))
    
    return empresas


def envia_experiencia(comp):
    # enquanto a janela de enviar não aparece clica no botão de enviar
    while not _find_img('publicar_doc.png', conf=0.9):
        _click_img('enviar_arquivo.png', conf=0.9)
        time.sleep(2)
    
    # verifica a pasta que está selecionada para inserir o arquivo no sistema
    # if not _find_img('pasta_pessoal_outros.png', conf=0.9):
    _click_img('drop.png', conf=0.9)
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
        
    p.press('tab')
    time.sleep(0.5)
    p.write(f'Relação de Empregados - Contratos_Vencimento_Modelo_Veiga - {comp}')
        
    # confirma o envio do arquivo
    p.hotkey('Alt', 'g')
    time.sleep(5)


def define_data():
    # Obter a data atual
    hoje = date_time.date.today()

    # Subtrair 90 dias da data atual
    data_subtraida = hoje - date_time.timedelta(days=90)

    hoje = hoje.strftime('%d/%m/%Y')
    data_subtraida = data_subtraida.strftime('%d/%m/%Y')

    return hoje.replace('//', '/'), data_subtraida.replace('//', '/')


@_time_execution
def run():
    # define o nome da planilha de andamentos
    dia = datetime.now().day
    mes = datetime.now().month
    ano = datetime.now().year
    
    comp = f'{dia}-{mes}-{ano}'
    
    andamentos = 'Experiência a Vencer'
    # pergunta se deve gerar uma nova planilha de dados
    novo = p.confirm(title='Script incrível', text='Gerar nova planilha de dados?', buttons=('Sim', 'Não'))

    # _login_web()
    # _abrir_modulo('folha')
    
    # se não for gerar uma nova planilha, seleciona a que já existe e pergunta se vai continuar de onde parou
    if novo == 'Não':
        empresas = _open_lista_dados()
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is None:
            return False
    # gera uma nova planilha e a seleciona
    else:
        index = 0
        gera_arquivo(comp, andamentos)
        empresas = pega_empresas_com_exp()
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa)
        
        # abre a empresa no domínio
        if not _login(empresa, andamentos):
            continue
        # gera o arquivo específico da empresa
        gera_arquivo(comp, andamentos, cod=empresa[0], cnpj=empresa[1], nome=empresa[2])


if __name__ == '__main__':
    run()
