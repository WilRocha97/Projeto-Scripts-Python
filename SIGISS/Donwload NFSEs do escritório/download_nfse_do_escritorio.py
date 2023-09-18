# -*- coding: utf-8 -*-
import time, fitz, re, os, shutil, xlrd, openpyxl
from bs4 import BeautifulSoup
from requests import Session
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from pyautogui import alert, confirm


def open_lista_dados():
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title='Abrir planilha com as chaves de acesso das notas',
        initialdir='ignore',
        filetypes=[('Excel','*.xlsx *.xls *.csv')]
    )
    
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
    print('✔ Nota Fiscal salva')


def indice(count, total_empresas, empresa, index=0):
    if count > 1:
        print(f'[ {len(total_empresas) - (count - 1)} Restantes ]\n\n')
    # Cria um indice para saber qual linha dos dados está
    indice_dados = f'[ {str(count + index)} de {str(len(total_empresas) + index)} ]'
    
    empresa = str(empresa).replace("('", '[ ').replace("')", ' ]').replace("',)", " ]").replace(',)', ' ]').replace("', '", ' - ')
    
    print(f'{indice_dados} - {empresa}')
    
            
def download_nota(count, total_empresas, nota, s, pasta_final):
    # printa o indice da nota que está sendo baixada
    indice(count, total_empresas, nota)
    chave = str(nota)
    
    # verifica se a chave de acesso não contem letras, se contem anota e retorna para executar a próxima
    if re.search(r'[a-zA-Z]', chave):
        escreve_relatorio_csv(f"'{nota};Não é uma chave válida", nome='Ocorrências', local=pasta_final)
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
        escreve_relatorio_csv(f"'{chave};O Código de Autenticidade da Nota Fiscal eletrônica informado é inválido.", nome='Ocorrências', local=pasta_final)
        print(f'❌ O Código de Autenticidade da Nota Fiscal eletrônica informado é inválido.')
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
    print('>>> Analisando a nota fiscal')
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
                
                dados_notas = ';'.join([cnpj_cliente,
                                                 nome_cliente,
                                                 uf_cliente,
                                                 municipio_cliente,
                                                 endereco_cliente,
                                                 numero_nota,
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
                
                cnpj_escritorio = re.compile(r'CNPJ : (\d\d\.\d\d\d\.\d\d\d/\d\d\d\d-\d\d)').search(textinho).group(1)
                # cria arquivo .txt para importar no dóminio web
                cria_txt(cnpj_escritorio, dados_notas, pasta_final)
                escreve_relatorio_csv(dados_notas, nome='Dados das notas', local=pasta_final)
            
            except():
                print(textinho)
    # renomeia a nota colocando o número dela
    shutil.move(arq, os.path.join(pasta_final, 'Notas', f'nfe_{numero_nota}.pdf'))
    
    return

def cria_txt(cnpj, dados_nota, pasta_final, tipo='NFSe', encode='latin-1'):
    cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
    codigos_dominio = {'26973312000175': '2945',
                       '26973312000256': '2950',
                       '26973312000337': '3364',
                       '26973312000418': '3387'}
    
    # cria um txt com os dados das notas em uma pasta nomeada com o código no domínio e o cnpj do prestador
    local = os.path.join(pasta_final, 'Arquivos para Importação', str(codigos_dominio[cnpj]) + '-' + str(cnpj))
    os.makedirs(local, exist_ok=True)
    
    try:
        f = open(os.path.join(local, f"{tipo} - {cnpj}.txt"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{tipo}  - {cnpj} - auxiliar.txt"), 'a', encoding=encode)
    
    f.write(str(dados_nota) + '\n')
    f.close()
    

def run():
    arquivo, notas, tipo_dados = open_lista_dados()
    pasta_final = ask_for_dir()
    # Abre o site do SIGISS pronto para validar as notas
    with Session() as s:
        s.get('https://valinhos.sigissweb.com/validarnfe')
        
        alert(title='Download NFSe do Escritório', text=f'Planilha selecionada: {arquivo.split("/")[-1]}\n\n'
                                                        f'Os resultados serão salvos na pasta "Notas Fiscais de Serviço"'
                                                        f' dentro do seguinte diretório: {pasta_final.replace("/", " >> ")}')
        
        pasta_final = os.path.join(pasta_final, 'Notas Fiscais de Serviço')
        quantidade_notas = 0
        if tipo_dados == 'csv':
            # cria o indice para cada empresa da lista de dados
            total_empresas = notas
            for count, nota in enumerate(notas, start=1):
                nota = nota.replace('\n', '')
                if download_nota(count, total_empresas, nota, s, pasta_final):
                    quantidade_notas += 1
        
        elif tipo_dados == 'xls':
            # cria o indice para cada empresa da lista de dados
            total_empresas = range(notas.nrows)
            for count, nota in enumerate(range(notas.nrows), start=1):
                # pega a última coluna da planilha
                nota = notas.cell_value(nota, -1)
                if download_nota(count, total_empresas, nota, s, pasta_final):
                    quantidade_notas += 1
                
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
                    
        alert(title='Download NFSe do Escritório', text=f'Download concluído, total de notas: {quantidade_notas}')
        
        s.close()
    
    
if __name__ == '__main__':
    run()
