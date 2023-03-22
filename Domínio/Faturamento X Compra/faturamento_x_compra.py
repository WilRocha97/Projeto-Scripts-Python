# -*- coding: utf-8 -*-
import fitz, re, shutil, pyperclip, time, os, pyautogui as p
from datetime import datetime
from dateutil.relativedelta import relativedelta
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _escreve_header_csv, e_dir, _open_lista_dados, _where_to_start, ask_for_dir
from dominio_comum import _login_web, _abrir_modulo, _login, _salvar_pdf


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
        time.sleep(1)
    
    time.sleep(1)
    p.write(f'012022')
    
    time.sleep(1)
    p.press('tab', presses=2)
    
    time.sleep(1)
    p.write(f'122022')
    
    # gera o relatório
    time.sleep(1)
    p.hotkey('alt', 'o')
    
    if not espera_gerar(empresa, andamento):
        return False

    guia = os.path.join('C:', 'Demonstrativo Mensal.pdf')
    while not os.path.exists(guia):
        print('>>> Aguardando salvar')
        if not _salvar_pdf():
            p.hotkey('alt', 'o')
            
            if not espera_gerar(empresa, andamento):
                return False
    
    p.press('esc', presses=4)
    time.sleep(2)
    
    arquivo = mover_demonstrativo(empresa, ano)
    
    captura_info_pdf(andamento, arquivo)
    
    
    print('✔ Demonstrativo Mensal gerado')


def espera_gerar(empresa, andamento):
    cod, cnpj, nome = empresa
    timer = 0
    # espera gerar
    while not _find_img('demonstrativo_mensal_gerado.png', conf=0.9):
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
                meses = re.compile(r'(\w.+)\n(.+,.+)\n(.+,.+)\n(.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)\n(.+,.+)').findall(textinho)
                totais = re.compile(r'Totais\n.+\n(.+)\n(.+)\n(.+)').search(textinho)
                
                totais_entrada = totais.group(1)
                totais_saida = totais.group(2)
                totais_servico = totais.group(3)

                entrada_jan ,saida_jan, servico_jan= 0, 0, 0
                entrada_fev, saida_fev, servico_fev = 0, 0, 0
                entrada_mar, saida_mar, servico_mar = 0, 0, 0
                entrada_abr, saida_abr, servico_abr = 0, 0, 0
                entrada_mai, saida_mai, servico_mai = 0, 0, 0
                entrada_jun, saida_jun, servico_jun = 0, 0, 0
                entrada_jul, saida_jul, servico_jul = 0, 0, 0
                entrada_ago, saida_ago, servico_ago = 0, 0, 0
                entrada_set, saida_set, servico_set = 0, 0, 0
                entrada_out, saida_out, servico_out = 0, 0, 0
                entrada_nov, saida_nov, servico_nov = 0, 0, 0
                entrada_dez, saida_dez, servico_dez = 0, 0, 0
                
                for mes in meses:
                    # Procura o valor a recolher da empresa
                    competencia = mes[0]
                    ano = mes[3]
                    entrada = mes[4]
                    saida = mes[5]
                    servico = mes[6]
                    
                    print(f'{competencia}/{ano} - Entrada: {entrada} - Saída: {saida} - Serviço: {servico}')
                    
                    if competencia == 'Janeiro':
                        entrada_jan = entrada
                        saida_jan = saida
                        servico_jan = servico
                    
                    if competencia == 'Fevereiro':
                        entrada_fev = entrada
                        saida_fev = saida
                        servico_fev = servico
                    
                    if competencia == 'Março':
                        entrada_mar = entrada
                        saida_mar = saida
                        servico_mar = servico
                    
                    if competencia == 'Abril':
                        entrada_abr = entrada
                        saida_abr = saida
                        servico_abr = servico
                    
                    if competencia == 'Maio':
                        entrada_mai = entrada
                        saida_mai = saida
                        servico_mai = servico

                    if competencia == 'Junho':
                        entrada_jun = entrada
                        saida_jun = saida
                        servico_jun = servico
                        
                    if competencia == 'Julho':
                        entrada_jul = entrada
                        saida_jul = saida
                        servico_jul = servico
                    
                    if competencia == 'Agosto':
                        entrada_ago = entrada
                        saida_ago = saida
                        servico_ago = servico
                    
                    if competencia == 'Setembro':
                        entrada_set = entrada
                        saida_set = saida
                        servico_set = servico
                        
                    if competencia == 'Outubro':
                        entrada_out = entrada
                        saida_out = saida
                        servico_out = servico
                        
                    if competencia == 'Novembro':
                        entrada_nov = entrada
                        saida_nov = saida
                        servico_nov = servico
                        
                    if competencia == 'Dezembro':
                        entrada_dez = entrada
                        saida_dez = saida
                        servico_dez = servico

                print(f'{competencia}/{ano} - Total: {totais_entrada} - Total: {totais_saida} - Total: {totais_servico}')
                
                _escreve_relatorio_csv(f"{cod};{cnpj};{nome};Entrada;"
                                       f"{entrada_jan};{entrada_fev};{entrada_mar};"
                                       f"{entrada_abr};{entrada_mai};{entrada_jun};"
                                       f"{entrada_jul};{entrada_ago};{entrada_set};"
                                       f"{entrada_out};{entrada_nov};{entrada_dez};{totais_entrada}", nome=andamento)

                _escreve_relatorio_csv(f"{cod};{cnpj};{nome};Saída;"
                                       f"{saida_jan};{saida_fev};{saida_mar};"
                                       f"{saida_abr};{saida_mai};{saida_jun};"
                                       f"{saida_jul};{saida_ago};{saida_set};"
                                       f"{saida_out};{saida_nov};{saida_dez};{totais_saida}", nome=andamento)

                _escreve_relatorio_csv(f"{cod};{cnpj};{nome};Serviço;"
                                       f"{servico_jan};{servico_fev};{servico_mar};"
                                       f"{servico_abr};{servico_mai};{servico_jun};"
                                       f"{servico_jul};{servico_ago};{servico_set};"
                                       f"{servico_out};{servico_nov};{servico_dez};{totais_servico}", nome=andamento)
            
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

        # _login_web()
        # _abrir_modulo('escrita_fiscal')

        for count, empresa in enumerate(empresas[index:], start=1):
            _indice(count, total_empresas, empresa)
            
            if not _login(empresa, andamentos):
                continue
            faturamento_compra(str(ano), empresa, andamentos)
    
    _escreve_header_csv('CÓDIGO;CNPJ;NOME;SITUAÇÃO;JANEIRO;FEVEREIRO;MARÇO;ABRIL;MAIO;JUNHO;JULHO;AGOSTO;SETEMBRO;OUTUBRO;NOVEMBRO;DEZEMBRO;TOTAIS', nome=andamentos)


if __name__ == '__main__':
    run()
