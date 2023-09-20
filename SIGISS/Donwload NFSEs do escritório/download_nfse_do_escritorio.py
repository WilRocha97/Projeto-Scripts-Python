# -*- coding: utf-8 -*-
import time, fitz, re, os, shutil, xlrd, openpyxl, threading, pandas as pd, PySimpleGUI as sg
from pathlib import Path
from bs4 import BeautifulSoup
from requests import Session
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from pyautogui import alert, confirm


def open_lista_dados(input_excel):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = input_excel
    
    if not file:
        return False
    
    try:
        # abre se for .xls
        if file.endswith('.xls') or file.endswith('.XLS'):
            workbook = xlrd.open_workbook(file)
            workbook = workbook.sheet_by_index(0)
            tipo_dados = 'xls'
        
        # abre se for .xlsx
        elif file.endswith('.xlsx') or file.endswith('.XLSX'):
            workbook = openpyxl.load_workbook(file)
            workbook = workbook['Plan1']
            tipo_dados = 'xlsx'
        
        # abre se for .csv
        elif file.endswith('.csv') or file.endswith('.CSV'):
            with open(file, 'r', encoding='latin-1') as f:
                workbook = f.readlines()
            tipo_dados = 'csv'
    
    # abre um alerta se não conseguir abrir o arquivo
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    return file, workbook, tipo_dados


def ask_for_dir(title='Selecione a pasta onde os arquivos serão salvos'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    # pergunta onde irá salvar os arquivos gerados pelo script
    folder = askdirectory(
        title=title,
    )
    
    return folder if folder else False


# Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'
def escreve_relatorio_csv(texto, nome='resumo', local='', end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()
    
    
def salvar_arquivo(nome, response, pasta_final):
    os.makedirs(os.path.join(pasta_final, 'Notas'), exist_ok=True)
    arquivo = open(os.path.join(pasta_final, 'Notas', nome), 'wb')
    
    # salva o PDF baseado na resposta da requisição com a chave da nota
    for parte in response.iter_content(100000):
        arquivo.write(parte)
    arquivo.close()
    # print('✔ Nota Fiscal salva')


def indice(count, total_empresas, empresa, index=0):
    """if count > 1:
        print(f'[ {len(total_empresas) - (count - 1)} Restantes ]\n\n')"""
    # Cria um indice para saber qual linha dos dados está
    indice_dados = f'[ {str(count + index)} de {str(len(total_empresas) + index)} ]'
    
    empresa = str(empresa).replace("('", '[ ').replace("')", ' ]').replace("',)", " ]").replace(',)', ' ]').replace("', '", ' - ')
    
    # print(f'{indice_dados} - {empresa}')
    
            
def download_nota(count, total_empresas, nota, s, pasta_final):
    # printa o indice da nota que está sendo baixada
    indice(count, total_empresas, nota)
    chave = str(nota)
    
    # verifica se a chave de acesso não contem letras, se contem anota e retorna para executar a próxima
    if re.search(r'[a-zA-Z]', chave):
        escreve_relatorio_csv(f"'{nota};Não é uma chave válida", nome='Andamentos', local=pasta_final)
        return False
    
    query = {'codigo': chave,
             'operacao': 'validar'}

    # faz a validação das notas
    response = s.post('https://valinhos.sigissweb.com/nfecentral?oper=validanfe&codigo={}&tipo=V'.format(chave), data=query)
    # define o nome do PDF da nota
    nome_nota = 'nfe_' + chave + '.pdf'
    soup = BeautifulSoup(response.content, 'html.parser', from_encoding="iso-8859-1")
    soup = soup.prettify()
    
    # verifica se o número da chave é uma chave válida, se não, anota e retorna para executar a próxima
    if re.compile(r'O Código de Autenticidade da Nota Fiscal eletrônica informado é inválido.').search(soup):
        escreve_relatorio_csv(f"'{chave};O Código de Autenticidade da Nota Fiscal eletrônica informado é inválido.", nome='Andamentos', local=pasta_final)
        # print(f'❌ O Código de Autenticidade da Nota Fiscal eletrônica informado é inválido.')
        s.get('https://valinhos.sigissweb.com/validarnfe')
        return False
    
    # salva o PDF da nota
    salvar_arquivo(nome_nota, response, pasta_final)
    time.sleep(0.5)
    # analisa a nota e coleta informações
    analisa_nota(nome_nota, pasta_final)
    
    return True
        

def analisa_nota(nome_nota, pasta_final):
    # Abrir o pdf
    arq = os.path.join(pasta_final, 'Notas', nome_nota)
    # print('>>> Analisando a nota fiscal')
    with fitz.open(arq) as pdf:
        # Para cada página do pdf
        for count, page in enumerate(pdf):
            try:
                # Pega o texto da pagina
                textinho = page.get_text('text', flags=1 + 2 + 8)
                # print(textinho)
                # time.sleep(12)
                try:
                    cnpj_cliente = re.compile(r'(\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)\nTELEFONE / FAX').search(textinho).group(1)
                except:
                    try:
                        cnpj_cliente = re.compile(r'(\d\d\d\.\d\d\d\.\d\d\d-\d\d)\nTELEFONE / FAX').search(textinho).group(1)
                    except:
                        cnpj_cliente = ''
                cnpj_cliente = cnpj_cliente.replace('.', '').replace('/', '').replace('-', '')
                
                nome_cliente = re.compile(r'DATA DO CADASTRO.+\n(.+)').search(textinho).group(1)
                if nome_cliente == 'IMPOSTOS FEDERAIS RETIDOS':
                    nome_cliente = re.compile(r'DATA DO CADASTRO(.+\n){5}(.+)').search(textinho).group(2)
                
                try:
                    uf_cliente = re.compile(r'UF\n(.+)').search(textinho).group(1)
                except:
                    uf_cliente = ''
                    
                try:
                    municipio_cliente = re.compile(r'MUNICÍPIO\n.+\n.+\n(.+)').search(textinho).group(1)
                except:
                    municipio_cliente = ''
                
                endereco_cliente = re.compile(r'(.+)\n.+\nTELEFONE / FAX').search(textinho).group(1)
                if endereco_cliente == 'INSCRIÇÃO ESTADUAL':
                    endereco_cliente = re.compile(r'(.+)\nTELEFONE / FAX').search(textinho).group(1)
                    
                numero_nota = re.compile(r'(.+)\nNÚMERO').search(textinho).group(1)
                serie_nota = re.compile(r'ELETRÔNICA DE\nSERVIÇO\n.+\n(.+)').search(textinho).group(1)
                data_nota = re.compile(r'(.+)\n.+\nDATA EMISSÃO').search(textinho).group(1)
                
                # se a nota for válida a situação é 0 se for cancelada é 2
                situacao = re.compile(r'DOCUMENTO VALIDADO COM SUCESSO').search(textinho)
                if situacao:
                    situacao = '0'
                else:
                    situacao = '2'
                    
                acumulador = '6102'
                
                # se o município do tomador for valinhos, o CFPS é 9101, se não é 9102
                if municipio_cliente == 'Valinhos':
                    cfps = '9101'
                else:
                    cfps = '9102'
                    
                valor_servico = re.compile(r'VALOR BRUTO DA NOTA FISCAL\n.+\n.+\nR\$ (.+)').search(textinho).group(1)
                valor_descontos = '0,00'
                valor_deducoes = re.compile(r'R\$ (.+)\n.+\n.+\n.+\nDEDUÇÕES').search(textinho).group(1)
                valor_contabil = valor_servico
                base_de_calculo = valor_servico
                aliquota_iss = re.compile(r'(\d,\d\d)(.+\n){7}ALIQUOTA ISS\(%\)').search(textinho).group(1)
                
                # se o ISS não for retido o valor fica no campo do ISS, se for retido, ele fica no campo do ISS retido
                if re.compile(r'O ISS NÃO DEVE SER RETIDO'):
                    valor_iss = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                    iss_retido = '0,00'
                else:
                    valor_iss = '0,00'
                    iss_retido = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                
                # tenta pegar o valor dos impostos a seguir, se não conseguir, coloca '0,00' ná variável
                try:
                    valor_irrf = re.compile(r'R\$ (.+)\n.+\nIRRF').search(textinho).group(1)
                except:
                    valor_irrf = '0,00'
                try:
                    valor_pis = re.compile(r'R\$ (.+)\n.+\nPIS').search(textinho).group(1)
                except:
                    valor_pis = '0,00'
                try:
                    valor_cofins = re.compile(r'R\$ (.+)\n.+\nCOFINS').search(textinho).group(1)
                except:
                    valor_cofins = '0,00'
                try:
                    valor_csll = re.compile(r'R\$ (.+)\n.+\nCSLL').search(textinho).group(1)
                except:
                    valor_csll = '0,00'
                
                valor_crf = '0,00'
                valor_inss = '0,00'
                codigo_do_iten = '17.18'
                quantidade = '1,00'
                valor_unitario = valor_servico
                
                dados_notas_cliente = ';'.join([cnpj_cliente, nome_cliente, uf_cliente, municipio_cliente, endereco_cliente])
                dados_notas_escritorio = ';'.join(['26.973.312/0001-75',
                                        'R. POSTAL SERVICOS CONTABEIS LTDA',
                                        'SP',
                                        'Valinhos',
                                        'Rua Fioravante Basílio Maglio, nº 258',])
                dados_notas = ';'.join([numero_nota,
                                        serie_nota,
                                        data_nota,
                                        situacao,
                                        acumulador,
                                        cfps,
                                        valor_servico,
                                        valor_descontos,
                                        valor_deducoes,
                                        valor_contabil,
                                        base_de_calculo,
                                        aliquota_iss,
                                        valor_iss,
                                        iss_retido,
                                        valor_irrf,
                                        valor_pis,
                                        valor_cofins,
                                        valor_csll,
                                        valor_crf,
                                        valor_inss,
                                        codigo_do_iten,
                                        quantidade,
                                        valor_unitario])
                
                pasta_final_nota = os.path.join(pasta_final, 'Notas')
                nome_nota_final = f'nfe_{numero_nota}.pdf'
                
                # se for CPF, anota ocorrência, renomeia a nota com o CPF do tomador e move para uma pasta separada
                if 0 < len(cnpj_cliente) < 14:
                    escreve_relatorio_csv(f"NFSe_{numero_nota}.pdf;Tomador com CPF, não gera arquivo para importar;'{cnpj_cliente}", nome='Andamentos', local=pasta_final)
                    pasta_final_nota = os.path.join(pasta_final, 'Notas CPF')
                    nome_nota_final = f'nfe_{numero_nota}_tomador_{cnpj_cliente}.pdf'
                    os.makedirs(pasta_final_nota, exist_ok=True)
                    
                # se não constar CNPJ do cliente na nota, anota a ocorrência
                elif cnpj_cliente == '':
                    escreve_relatorio_csv(f"NFSe_{numero_nota}.pdf;Tomador sem CNPJ informado", nome='Andamentos', local=pasta_final)
                
                # se tiver busca o código do Domínio na planilha
                else:
                    codigo_cliente = busca_codigo_cliente(cnpj_cliente)
                    # se não encontrar o CNPJ do cliente na planilha de códigos, anota a ocorrência
                    if codigo_cliente == 'Cliente não encontrado na planilha de códigos':
                        escreve_relatorio_csv(f"NFSe_{numero_nota}.pdf;{codigo_cliente};'{cnpj_cliente}", nome='Andamentos', local=pasta_final)
                    elif codigo_cliente == 'CÓDIGO DO DOMÍNIO':
                        escreve_relatorio_csv(f"NFSe_{numero_nota}.pdf;CNPJ do cliente repetido na planilha de códigos, verificar qual consta com o código correto e deletar os demais;'{cnpj_cliente}", nome='Andamentos', local=pasta_final)
                    # se o CNPJ do cliente for encontrado na planilha de códigos, cria arquivo .txt para importar no dóminio web
                    else:
                        cria_txt(codigo_cliente, cnpj_cliente, f'{dados_notas_escritorio};{dados_notas}', pasta_final)
                        escreve_relatorio_csv(f"NFSe_{numero_nota}.pdf;Arquivo para importação criado com sucesso;{codigo_cliente} - {cnpj_cliente}", nome='Andamentos', local=pasta_final)
                        
                escreve_relatorio_csv(f'{dados_notas_escritorio};{dados_notas_cliente};{dados_notas}', nome='Dados das notas', local=pasta_final)
                
            
            except():
                print(textinho)
    # renomeia a nota colocando o número dela
    shutil.move(arq, os.path.join(pasta_final_nota, nome_nota_final))
    return


def cria_txt(codigo_cliente, cnpj_cliente, dados_nota, pasta_final, tipo='NFSe', encode='latin-1'):
    # cria um txt com os dados das notas em uma pasta nomeada com o código no domínio e o cnpj do prestador
    local = os.path.join(pasta_final, 'Arquivos para Importação', str(codigo_cliente) + '-' + str(cnpj_cliente))
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{tipo} - {cnpj_cliente}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{tipo}  - {cnpj_cliente} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(dados_nota) + '\n')
    f.close()
    

def busca_codigo_cliente(cnpj_cliente):
    try:
        # print('>>> Consultando código do cliente')
        # Definir o nome das colunas
        colunas = ['cnpj', 'numero', 'nome']
        # Localiza a planilha
        caminho = Path('Dados_clientes/códigos clientes.xlsx')
        # Abrir a planilha
        lista = pd.read_excel(caminho)
        # Definir o index da planilha
        lista.set_index('CNPJ', inplace=True)
        
        cliente = lista.loc[int(cnpj_cliente)]
        cliente = list(cliente)
        codigo_tomador = cliente[0]
        
        # print(f'{cnpj_cliente} - {codigo_tomador}')
        return codigo_tomador
    except:
        # print(f' {cnpj_cliente} - Cliente não encontrado na planilha de códigos')
        return 'Cliente não encontrado na planilha de códigos'
    

def run(window, input_excel, output_dir):
    arquivo, notas, tipo_dados = open_lista_dados(input_excel)
    pasta_final = output_dir
    # Abre o site do SIGISS pronto para validar as notas
    with Session() as s:
        s.get('https://valinhos.sigissweb.com/validarnfe')
        
        pasta_final = os.path.join(pasta_final, 'Notas Fiscais de Serviço')
        quantidade_notas = 0
        if tipo_dados == 'csv':
            # cria o indice para cada empresa da lista de dados
            total_empresas = notas
            for count, nota in enumerate(notas, start=1):
                nota = nota.replace('\n', '')
                if download_nota(count, total_empresas, nota, s, pasta_final):
                    quantidade_notas += 1
                window['progressbar'].update_bar(count, max=int(len(notas)))
                window['Progresso_texto'].update(str( int( int(count) / int(len(notas)) *100 ) ) + '%')
                window.refresh()

                # Verifica se o usuário solicitou o encerramento do script
                if event == 'Encerrar':
                    return
        
        elif tipo_dados == 'xls':
            # cria o indice para cada empresa da lista de dados
            total_empresas = range(notas.nrows)
            for count, nota in enumerate(range(notas.nrows), start=1):
                # pega a última coluna da planilha
                nota = notas.cell_value(nota, -1)
                if download_nota(count, total_empresas, nota, s, pasta_final):
                    quantidade_notas += 1
                window['progressbar'].update_bar(count, max=int(total_empresas))
                window['Progresso_texto'].update(str( int( int(count) / int(total_empresas) *100 ) ) + '%')
                window.refresh()
                
                # Verifica se o usuário solicitou o encerramento do script
                if event == 'Encerrar':
                    return
                
        elif tipo_dados == 'xlsx':
            last_column = notas.max_column
            # cria o indice para cada empresa da lista de dados
            total_empresas = []
            for nota in notas.iter_rows(min_row=1, min_col=last_column, max_col=last_column):
                total_empresas.append(nota)
            
            for count, nota in enumerate(notas.iter_rows(min_row=1, min_col=last_column, max_col=last_column), start=1):
                for chave in nota:
                    if download_nota(count, total_empresas, chave.value, s, pasta_final):
                        quantidade_notas += 1
                    window['progressbar'].update_bar(count, max=int(len(total_empresas)))
                    window['Progresso_texto'].update(str( int( int(count) / int(len(total_empresas)) *100 ) ) + '%')
                    window.refresh()

                    # Verifica se o usuário solicitou o encerramento do script
                    if event == 'Encerrar':
                        return
    
        s.close()
        alert(text=f'Download concluído, {quantidade_notas} notas salvas')
    

# Define o ícone global da aplicação
sg.set_global_icon('Assets/auto-flash.ico')
if __name__ == '__main__':
    sg.theme('GrayGrayGray')  # Define o tema do PySimpleGUI
    # sg.theme_previewer()
    # Layout da janela
    layout = [
        [sg.Button('Ajuda'), sg.Button('Lista de códigos do Domínio')],
        [sg.Text('Selecione um arquivo Excel com as chaves de acesso das notas:')],
        [sg.InputText(key='input_excel'), sg.FileBrowse(key='Abrir', file_types=(("Planilhas Excel", "*.xlsx"),))],
        [sg.Text('Selecione um diretório para salvar os resultados:')],
        [sg.InputText(key='output_dir'), sg.FolderBrowse(key='Abrir2')],
        [sg.Button('Iniciar'), sg.Button('Encerrar', disabled=True), sg.Button('Abrir resultados', disabled=True)],
        [sg.Text('', key='Progresso_texto')],
        [sg.ProgressBar(max_value=0, orientation='h', size=(50, 15), key='progressbar', bar_color=('#fca400', '#d4d4d4'))],
    ]
    
    window = sg.Window('Download NFSE_VP SIGISSWEB', layout)
    
    def run_script_thread():
        
        
        if not input_excel or not output_dir:
            alert(text=f'Por favor selecione uma planilha e um diretório.')
            return
        
        window['input_excel'].update(disabled=True)
        window['output_dir'].update(disabled=True)
        window['Abrir'].update(disabled=True)
        window['Abrir2'].update(disabled=True)
        window['Iniciar'].update(disabled=True)
        window['Encerrar'].update(disabled=False)
        
        # Chama a função que executa o script
        run(window, input_excel, output_dir)
        
        window['input_excel'].update(disabled=False)
        window['output_dir'].update(disabled=False)
        window['Abrir'].update(disabled=False)
        window['Abrir2'].update(disabled=False)
        window['Iniciar'].update(disabled=False)
        window['Encerrar'].update(disabled=True)
        window['Abrir resultados'].update(disabled=False)
    
    
    while True:
        event, values = window.read()
        try:
            input_excel = values['input_excel']
            output_dir = values['output_dir']
        except:
            output_dir = 'Desktop'
            pass
        
        if event == sg.WIN_CLOSED:
            break
        
        elif event == 'Ajuda':
            os.startfile('Manual do usuário - Download NFSe_VP SIGISSWEB.docx')
          
        elif event == 'Lista de códigos do Domínio':
            os.startfile(os.path.join('Dados_clientes', 'códigos clientes.xlsx'))
            
        elif event == 'Iniciar':
            # Cria uma nova thread para executar o script
            script_thread = threading.Thread(target=run_script_thread)
            script_thread.start()
            
        elif event == 'Encerrar':
            window['Encerrar'].update(disabled=True)
            alert(text='Download encerrado.\n\n'
                     'Caso queira reiniciar o download, apague os arquivos gerados anteriormente ou selecione um novo local.\n\n'
                     'O Script não continua uma execução já iniciada.\n\n')
            window['Iniciar'].update(disabled=False)
            window['Abrir resultados'].update(disabled=False)
            window['input_excel'].update(disabled=False)
            window['output_dir'].update(disabled=False)
            
        elif event == 'Abrir resultados':
            os.startfile(os.path.join(output_dir, 'Notas Fiscais de Serviço'))
            
    if event == 'Iniciar':
        window.close()
