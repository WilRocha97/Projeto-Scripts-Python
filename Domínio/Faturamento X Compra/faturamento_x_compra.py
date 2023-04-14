# -*- coding: utf-8 -*-
import fitz, re, shutil, pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, e_dir, _open_lista_dados, _where_to_start, ask_for_dir
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf, _verifica_dominio, _encerra_dominio


def faturamento_compra(ano, empresa, andamento):
    _wait_img('relatorios.png', conf=0.9, timeout=-1)
    # Relatórios
    p.hotkey('alt', 'r')
    time.sleep(0.5)
    # Acompanhamentos
    p.press('a')
    # Demonstrativo Mensal
    time.sleep(0.5)
    p.press('m')
    # Confirmar
    time.sleep(0.5)
    p.press('enter')
    
    while not _find_img('demonstrativo_mensal.png', conf=0.9):
        verificacao = _verifica_dominio()
        if verificacao != 'continue':
            return verificacao
        time.sleep(1)
    
    time.sleep(1)
    p.write(f'01{ano}')
    
    time.sleep(1)
    p.press('tab', presses=2)
    
    time.sleep(1)
    p.write(f'12{ano}')
    
    # gera o relatório
    time.sleep(1)
    p.hotkey('alt', 'o')
    
    resultado = espera_gerar(empresa, andamento)
    if not resultado:
        return 'ok'
    elif resultado:
        pass
    else:
        return resultado

    guia = os.path.join('C:', 'Demonstrativo Mensal.pdf')
    while not os.path.exists(guia):
        print('>>> Aguardando salvar')
        if not _salvar_pdf():
            p.hotkey('alt', 'o')
            
            resultado = espera_gerar(empresa, andamento)
            if not resultado:
                return 'ok'
            elif resultado:
                pass
            else:
                return resultado
            
    p.press('esc', presses=4)
    time.sleep(2)
    
    arquivo = mover_demonstrativo(empresa, ano)
    captura_info_pdf(andamento, arquivo)
    
    print('✔ Demonstrativo Mensal gerado')
    return 'ok'


def espera_gerar(empresa, andamento):
    cod, cnpj, nome = empresa
    timer = 0
    # espera gerar
    while not _find_img('demonstrativo_mensal_gerado.png', conf=0.9):
        verificacao = _verifica_dominio()
        if verificacao != 'continue':
            return verificacao
        
        print('>>> Aguardando gerar')
        if _find_img('sem_dados.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Sem dados para imprimir']), nome=andamento)
            print('❌ Sem dados para imprimir')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=4, interval=1)
            time.sleep(1)
            return False
        time.sleep(1)
        timer += 1
        
        if timer >= 30:
            p.hotkey('alt', 'o')
            timer = 0
    
    return True


def mover_demonstrativo(empresa, ano):
    os.makedirs('execução/Demonstrativos', exist_ok=True)
    cod, cnpj, nome = empresa

    execucoes = 'V:\\Setor Robô\\Scripts Python\\Domínio\\Faturamento X Compra\\execução\\Demonstrativos'

    guia = os.path.join('C:', 'Demonstrativo Mensal.pdf')
    arquivo = f'{cod} - {nome.replace("/", " ")} - Demonstrativo Mensal {ano}.pdf'
    
    while not os.path.exists(guia):
        time.sleep(1)
    while os.path.exists(guia):
        try:
            shutil.move(guia, os.path.join(execucoes, arquivo))
            time.sleep(4)
        except:
            pass
    
    return arquivo
    

def captura_info_pdf(andamento, arquivo, refaz_planilha='não'):
    empresa = re.compile(r'(\d+) - (.+) - Demonstrativo Mensal').search(arquivo)
    cod = empresa.group(1)
    nome = empresa.group(2)
    
    if refaz_planilha == 'sim':
        arquivo = arquivo
    else:
        arquivo = os.path.join("V:\\Setor Robô\\Scripts Python\\Domínio\\Faturamento X Compra\\execução\\Demonstrativos", arquivo)
        
    with fitz.open(arquivo) as pdf:
        
        # Para cada página do pdf, se for a segunda página o script ignora
        for count, page in enumerate(pdf):
            if count == 1:
                continue
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                cnpj = re.compile(r'(.+)\nCNPJ:').search(textinho).group(1)
                # lista de meses com seus valores, o '.group' deve ser trocad pela posição na lista, exemplo: '.group(1) troca por 'item_da_lista[0]'
                meses = re.compile(r'(\w.+)\n.+,.+\n.+,.+\n.+\n(.+,.+)\n(.+,.+)\n(.+,.+)\n.+,.+\n.+,.+\n.+,.+\n.+,.+').findall(textinho)
                totais = re.compile(r'Totais\n.+\n(.+)\n(.+)\n(.+)').search(textinho)
                
                lista_meses  = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                          'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
                
                # para cada tipo, inicia a linha no '.csv' com as infos da empresa
                # para cada mês encontrado, insere os valores de todos os meses referêntes ao mesmo tipo para que eles fiquem na mesma linha
                # e os meses sejam separados por colunas
                # exemplo:
                # CÓDIGO | CNPJ | Nome |  TIPO   | Jan | Fev | Mar | Abr | Mai | Jun | jul | Ago | Sep | Out | Nov | Des | Totais
                # 000000 | 0000 | ABCD | Entrada | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0
                # 000000 | 0000 | ABCD |  Saída  | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0
                # 000000 | 0000 | ABCD | Serviço | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0 | 0,0
                
                # ()
                for tipo in [(1, 'Entrada'), (2, 'Saída'), (3, 'Serviço')]:
                    _escreve_relatorio_csv(f"{cod};{cnpj};{nome};{tipo[1]}", end=';', nome=andamento)
                    count = 0
                    for mes in meses:
                        # percorre uma lista de todos os meses em ondem verificando se o mes coletado no arquivo corresponde ao mesmo da lista
                        # se não adicona 0 na possição do mês não encontrado
                        while not mes[0] == lista_meses[count]:
                            _escreve_relatorio_csv('0,00', end=';', nome=andamento)
                            count += 1
                        else:
                            valor = mes[tipo[0]]
                            _escreve_relatorio_csv(f'{valor}', end=';', nome=andamento)
                            count += 1
                    
                    # apos terminar a lista de meses do arquivo e ainda não ter 12 valores inseridos na linha correspondente ao tipo,
                    # adiciona 0 na linha até completar 12 colunas e só assim insere o valor total do tipo
                    while count < 12:
                        _escreve_relatorio_csv('0,00', end=';', nome=andamento)
                        count += 1
                        
                    _escreve_relatorio_csv(f'{totais.group(tipo[0])}', nome=andamento)
                    print(f'✔ Valores referentes a {tipo[1]} coletados')
                    
            except():
                print(textinho)
                print('ERRO')


@_time_execution
def run():
    # configura o ano e digita no domínio
    ano = int(datetime.now().year)
    if int(datetime.now().month) < 3:
        ano -= 1
    andamentos = 'Faturamento X Compra - ' + str(ano)
    
    refaz_planilha = p.confirm(title='Script incrível', text='Gerar planilha com os demonstrativos já existentes?', buttons=['Sim', 'Não'])
    
    if refaz_planilha == 'Sim':
        documentos = ask_for_dir()
        # Analiza cada pdf que estiver na pasta
        for arq in os.listdir(documentos):
            print(f'\nArquivo: {arq}')
            # Abrir o pdf
            arq = os.path.join(documentos, arq)
            captura_info_pdf(andamentos, arq, refaz_planilha='sim')
        
    else:
        empresas = _open_lista_dados()
        
        index = _where_to_start(tuple(i[0] for i in empresas))
        if index is None:
            return andamentos
    
        total_empresas = empresas[index:]

        _login_web()
        _abrir_modulo('escrita_fiscal')

        for count, empresa in enumerate(empresas[index:], start=1):
            _indice(count, total_empresas, empresa)
            resultado = ''
            while resultado != 'ok':
                if not _login(empresa, andamentos):
                    break
            
                resultado = faturamento_compra(str(ano), empresa, andamentos)
                
                if resultado == 'dominio fechou':
                    _login_web()
                    _abrir_modulo('escrita_fiscal')
                    
                if resultado == 'modulo fechou':
                    _abrir_modulo('escrita_fiscal')
    
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;SITUAÇÃO;JANEIRO;FEVEREIRO;MARÇO;ABRIL;MAIO;JUNHO;JULHO;AGOSTO;SETEMBRO;OUTUBRO;NOVEMBRO;DEZEMBRO;TOTAIS', nome=andamentos)
    _encerra_dominio()

if __name__ == '__main__':
    _encerra_dominio()
    run()
