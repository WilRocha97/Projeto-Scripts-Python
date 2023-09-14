# -*- coding: utf-8 -*-
import time, fitz, re, os, shutil, xlrd, openpyxl
from bs4 import BeautifulSoup
from requests import Session
from tkinter.filedialog import askopenfilename, askdirectory, Tk
# from pyautogui import alert, confirm


def open_lista_dados(i_dir='ignore'):
    ftypes = [('Excel', '*.xlsx *.xls *.csv')]

    file = ask_for_file(filetypes=ftypes, initialdir=i_dir)
    if not file:
        return False

    try:
        # Abra o arquivo Excel
        try:
            workbook = xlrd.open_workbook(file)
            workbook = workbook.sheet_by_index(0)
        except:
            workbook = openpyxl.load_workbook(file)
            workbook = workbook['Plan1']
            
    except Exception as e:
        alert(title='Mensagem erro', text=f'Não pode abrir arquivo\n{str(e)}')
        return False

    return workbook


def ask_for_file(title='Abrir planilha com as chaves de acesso das notas', filetypes='*', initialdir=os.getcwd()):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
    file = askopenfilename(
        title=title,
        filetypes=filetypes,
        initialdir=initialdir
    )
    
    return file if file else False


def ask_for_dir(title='Abrir pasta de destino das notas'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    
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
    
    
def percorre_planilha(s, notas, pasta_final):
    try:
        # cria o indice para cada empresa da lista de dados
        total_empresas = range(notas.nrows)
        for count, nota in enumerate(range(notas.nrows)):
            nota = notas.cell_value(nota, -1)
            if count < 3:
                continue
            # printa o indice da empresa que está sendo executada
            indice(count, total_empresas, nota)
            download_nota(s, nota, pasta_final)
    except:
        last_column = notas.max_column
        # cria o indice para cada empresa da lista de dados
        total_empresas = []
        for nota in notas.iter_rows(min_row=1, min_col=last_column, max_col=last_column):
            total_empresas.append(nota)

        for count, nota in enumerate(notas.iter_rows(min_row=1, min_col=last_column, max_col=last_column)):
            for chave in nota:
                if count == 0:
                    continue
                # printa o indice da empresa que está sendo executada
                indice(count, total_empresas, chave.value)
                download_nota(s, chave.value, pasta_final)
            
            
def download_nota(s, nota, pasta_final):
        chave = str(nota)

        query = {'codigo': chave,
                 'operacao': 'validar'}

        # Faz a validação das notas
        response = s.post('https://valinhos.sigissweb.com/nfecentral?oper=validanfe&codigo={}&tipo=V'.format(chave), data=query)
        nome_nota = 'nfe_' + chave + '.pdf'
        # Salva o PDF da nota
        salvar_arquivo(nome_nota, response, pasta_final)
        # analisa a nota e coleta informações
        time.sleep(0.5)
        analisa_nota(nome_nota, pasta_final)
        

def analisa_nota(nome_nota, pasta_final):
    # Abrir o pdf
    arq = os.path.join(pasta_final, 'Notas', nome_nota)
    
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
                
                situacao = re.compile(r'DOCUMENTO VALIDADO COM SUCESSO').search(textinho)
                if situacao:
                    situacao = '0'
                else:
                    situacao = '2'
                    
                acumulador = '6102'
                
                if municipio_cliente == 'Valinhos':
                    cfps = '9101'
                else:
                    cfps = '9102'
                    
                valor_servico = re.compile(r'VALOR BRUTO DA NOTA FISCAL\n.+\n.+\nR\$ (.+)').search(textinho).group(1)
                valor_descontos = '0,00'
                valor_decucoes = re.compile(r'R\$ (.+)\n.+\n.+\n.+\nDEDUÇÕES').search(textinho).group(1)
                valor_contabil = valor_servico
                base_de_calculo = valor_servico
                aliquota_iss = re.compile(r'(\d,\d\d)(.+\n){7}ALIQUOTA ISS\(%\)').search(textinho).group(1)
                
                if re.compile(r'O ISS NÃO DEVE SER RETIDO'):
                    valor_iss = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                    iss_retido = '0,00'
                else:
                    valor_iss = '0,00'
                    iss_retido = re.compile(r'R\$ (.+)\n(.+\n){5}VALOR I.S.S.').search(textinho).group(1)
                    
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
                
                escreve_relatorio_csv(';'.join([cnpj_cliente,
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
                                                 valor_decucoes,
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
                                                 valor_unitario]),
                                       nome='Dados das notas', local=pasta_final)
            
            except():
                print(textinho)
    shutil.move(arq, os.path.join(pasta_final, 'Notas', f'nfe_{numero_nota}.pdf'))


def run():
    notas = open_lista_dados()
    pasta_final = ask_for_dir()
    # Abre o site do SIGISS pronto para validar as notas
    with Session() as s:
        s.get('https://valinhos.sigissweb.com/validarnfe')
        percorre_planilha(s, notas, os.path.join(pasta_final, 'Notas Fiscais de Serviço'))
        s.close()
    
    
if __name__ == '__main__':
    run()
