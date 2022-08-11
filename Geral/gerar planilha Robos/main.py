# -*- coding: utf-8 -*-
from tkinter.filedialog import askdirectory, askopenfilename, Tk
from pandas import read_excel, read_csv, notna
from datetime import datetime, date
import eel, os, re, sys
import fitz

eel.init('web', allowed_extensions=['.css', '.html', '.js', '.json'])

_RELATORIOS = '.'
_COMPLEMENTAR = '.'
_REL_PROC = ['a', 'b', 'c']
_CAMINHO = os.path.join(r'\\VPSRV02','DCA','Setor Robô', 'Relatorios')

_CONTABILISTAS = {
    r"^R\.? POSTAL|^RPEM|^RODRIGO": ('RPOSTAL','f7j54kymq4'),
    r"^RP |^RPMO|^EVANDRO": ('EVMASOL','ty63y26227'),
    r"^RCN|^ROSELEI": ('CNROSELEI','yds34b9p8k'),
    r"^AVJ|^ALDEMAR": ('VEIGAJR','yq4j8degy5'),
}

_CERTIFICADOS = {
    r'^R\.? POSTAL' : (r'..\certificados\CERT R POSTAL 26973312.pfx', '26973312'),
    r'^RP |^RPMO' : (r'..\certificados\CERT RPMO 05497487.pfx', '05497487'),
    r'^EVANDRO' : (r'..\certificados\CERT EVANDRO 307828.pfx', '307828'),
    r'^RODRIGO' : (r'..\certificados\CERT RODRIGO 273295.pfx', '273295'),
    r'^ROSELEI' : (r'..\certificados\CERT ROSELEI 045838.pfx', '045838'),
    r'^RPEM' : (r'..\certificados\CERT RPEM 35586086.pfx', '35586086'),
    r'^ALDEMAR' : (r'..\certificados\CERT VEIGA 251946.pfx', '251946'),
    r'^AVJ' : (r'..\certificados\CERT AVJ 19142855.pfx', '19142855'),
}


def plan_to_dataframe(origem):
    selected_arq = None
    caminho = os.path.join(_CAMINHO, origem)
    regexs = {
        'ae' : r'^ExpCli_(\d{4})(\d{2})(\d{2})\.[xlsXLS]{3}$',
        'dp' : r'^Relatorio DPCUCA (\d{2})(\d{2})(\d{4}).[csvCSV]{3}$',
        'pdf_dp': r'^dpcuca - (\d{2})(\d{2})(\d{4})\.[pdfPDF]{3}$'
    }
    
    prev_date = date(1970, 1, 1)

    for arq in os.listdir(caminho):
        match = re.search(regexs[origem], arq)
        if not match: continue
        
        comp = [int(i) for i in match.groups()]
        if origem != 'ae': comp.reverse()
        comp = date(*comp)

        if comp > prev_date:
            selected_arq = arq
            prev_date = comp

    if not selected_arq: return None

    print('Usando', selected_arq)
    ext = selected_arq[-4:].lower()
    if ext == '.csv':
        colunas = ['cod', 'cnpj', 'razao']
        return read_csv(os.path.join(caminho, selected_arq), header=None, names=colunas, sep=';', encoding='latin-1')
    if ext in ['.xls', 'xlsx']:
        return read_excel(os.path.join(caminho, selected_arq))
    if ext == '.pdf':
        return fitz.open(os.path.join(caminho, selected_arq))

    
def sub_proc(row):  
    cert, senha = ('', '')
    for regex, tupla in _CERTIFICADOS.items():
        if re.search(regex, str(row[2])):
            cert, senha = tupla
            break

    row[2] = cert
    row.append(senha)
    return row[1:]


def sub_fiscal(row):
    if len(row) < 5:
        for regex, tupla in _CONTABILISTAS.items():
            if re.search(regex, str(row[2])):
                usuario, senha = tupla
                break
        else:
            usuario, senha = ('', '')
        row[2] = usuario
        row.insert(3, senha)
    else:
        row[3] = row[3].replace("'","")

    return row[1:]


def salvar_py(nome, dados, modelo):
    file = open(nome, 'w')
    file.write('empresas = [\n')
    try:
        for row in dados:
            row = sub_proc(row) if modelo in _REL_PROC else sub_fiscal(row)
            row = [f'"{item}"' if item not in [1, 2] else str(item) for item in row]
            row = [item if '\\' not in item else 'r' + item for item in row]
            file.write('(' + ', '.join(row) + '),\n')
    except:
        file.close()
        return False

    file.write(']')
    file.close()
    return True


def salvar_csv(nome, dados, modelo, saida):
    file = open(nome, 'w')
    try:
        if saida == 'tratado':
            for row in dados:
                if row[-1] in [1, 2]: row.pop()
                row = sub_proc(row) if modelo in _REL_PROC else sub_fiscal(row)
                file.write(';'.join(row) + '\n')
        else:
            for row in dados:
                if row[-1] in [1, 2]: row.pop()
                file.write(';'.join(row) + '\n')
    except:
        file.close()
        return False

    file.close()
    return True


def gera_a(plan):
    dados = []
    conds = ['(Codigo < 20000)', '(Razao == Razao)','(CNPJ == CNPJ)', '(Certificado == Certificado)']
    plan.query(' and '.join(conds), inplace=True)
    for row in plan.itertuples():
        dados.append([row.Razao.replace('/', ''), str(row.CNPJ).replace('.0', '').rjust(14, '0'), row.Certificado])
    
    dados.sort(key=lambda item: item[2])
    return dados


def gera_b(plan):
    dados = []
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(CNPJ == CNPJ)', '(Certificado == Certificado)',
        'Classificacao.str.lower() == "simples nacional"'
    ]
    plan.query(' and '.join(conds), inplace=True)
    for row in plan.itertuples():
        dados.append([row.Razao.replace('/', ''), str(row.CNPJ).replace('.0', '').rjust(14, '0'), row.Certificado])

    dados.sort(key=lambda item: item[2])
    return dados


def gera_c(plan):
    dados = []
    filtro = ('lucro presumido', 'lucro real', 'inativa - lucro presumido', 'isentas/imunes', 'inativa - imune/isenta')
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(CNPJ == CNPJ)', '(Certificado == Certificado)',
        'Classificacao.str.lower() in @filtro'
    ]
    plan.query(' and '.join(conds), inplace=True)
    for row in plan.itertuples():
        dados.append([row.Razao.replace('/', ''), str(row.CNPJ).replace('.0', '').rjust(14, '0'), row.Certificado])

    dados.sort(key=lambda item: item[2])
    return dados


def gera_d(plan):
    dados_contribuinte = []
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(CNPJ == CNPJ)', '(Usuario == Usuario)', '(Senha == Senha)', 
        '(Contador.str.lower() == "contribuinte")'
    ]
    plan_contribuinte = plan.query(' and '.join(conds))
    for row in plan_contribuinte.itertuples():
        dados_contribuinte.append([row.Razao.replace('/', ''), str(row.CNPJ).replace('.0', '').rjust(14, '0'), row.Usuario, row.Senha, 1])

    dados_contador = []
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(CNPJ == CNPJ)', '(Contador == Contador)',
        '(Contador.str.lower() != "contribuinte")'
    ]
    plan_contador = plan.query(' and '.join(conds))
    for row in plan_contador.itertuples():
        dados_contador.append([row.Razao.replace('/', ''), str(row.CNPJ).replace('.0', '').rjust(14, '0'), row.Contador, 2])

    dados_contador.sort(key=lambda item: item[-2])
    dados_contribuinte.sort(key=lambda item: item[-3])
    dados = dados_contribuinte + dados_contador
    return dados


def gera_e(plan):
    dados_contribuinte = []
    filtro_ie = ('isento', 'baixado', 'baixada', 'inapto', 'none')
    filtro_classificacao = ('lucro presumido', 'lucro real', 'inativa - lucro presumido')
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(InsEstadual == InsEstadual)', '(Usuario == Usuario)', '(Senha == Senha)',
        '(Contador.str.lower() == "contribuinte")', 'InsEstadual.str.lower() not in @filtro_ie',
        'Classificacao.str.lower() in @filtro_classificacao'
    ]
    plan_contribuinte = plan.query(' and '.join(conds))
    for row in plan_contribuinte.itertuples():
        ie = ''.join(i for i in str(row.InsEstadual) if i.isdigit())
        if len(ie) > 12: continue
        dados_contribuinte.append([row.Razao.replace('/', ''), ie.rjust(12, '0'), row.Usuario, row.Senha, 1])

    dados_contador = []
    conds = [
        '(Codigo < 20000)', '(Razao == Razao)', '(InsEstadual == InsEstadual)', '(Contador == Contador)',
        '(Contador.str.lower() != "contribuinte")', 'InsEstadual.str.lower() not in @filtro_ie',
        'Classificacao.str.lower() in @filtro_classificacao'
    ]
    plan_contador = plan.query(' and '.join(conds))
    for row in plan_contador.itertuples():
        ie = ''.join(i for i in str(row.InsEstadual) if i.isdigit())
        if len(ie) > 12: continue
        dados_contador.append([row.Razao.replace('/', ''), ie.rjust(12, '0'), row.Contador, 2])

    dados_contador.sort(key=lambda item: item[-2])
    dados_contribuinte.sort(key=lambda item: item[-3])
    dados = dados_contribuinte + dados_contador
    return dados   


def gera_f(pdf):
    dados = []
    for page in pdf:
        texto = page.getText('text', flags=1+2+8).split('\n')

        for line in texto[5:]:
            if '|' in line: dados.append([item.strip() for item in line.split('|')])

    dados.sort(key=lambda item: item[-1])
    return dados

def gera_g(plan):
    dados_plan = []
    conds = ["(cod == cod)", "(cnpj == cnpj)", "(razao == razao)"]
    plan.query(" and ".join(conds))
    for row in plan.itertuples():
        cpf_cnpj = "".join([i for i in str(row.cnpj) if i.isdigit()])
        dados_plan.append([str(row.cod), str(row.razao).strip(), cpf_cnpj])

    colunas_extra = [
        'cod', 'cliente', 'prot', 'folha', 'c_905', 'c_115', 'IRF', 'PIS', 
        'sind', 'pensao', 'ret_11', 'sindical', 'RPA', 'prolabore', '',
    ]
    plan_extra = read_excel(_COMPLEMENTAR, names=colunas_extra, keep_default_na=False, header=1)
    plan_extra.query("c_905=='X'", inplace=True)
    dados_extra = [str(x.cod) for x in plan_extra.itertuples()]

    dados = [empresa for codigo in dados_extra for empresa in dados_plan if codigo == empresa[0]]
    dados.sort(key=lambda item: item[-2])
    return dados


def seleciona_modelo(modelo, df):
    if modelo == 'a':
        return gera_a(df)
    elif modelo == 'b':
        return gera_b(df)
    elif modelo == 'c':
        return gera_c(df)
    elif modelo == 'd':
        return gera_d(df)
    elif modelo == 'e':
        return gera_e(df)
    elif modelo == 'f':
        return gera_f(df)
    elif modelo == 'g':
        return gera_g(df)
    else:
        return False


@eel.expose
def run(args):
    comeco = datetime.now()
    print('>>> Começando a rotina, ', comeco)
    
    modelo, extensao, separacao, saida, origem = args
    print(args)
    
    df = plan_to_dataframe(origem)
    if df is None: return False

    #formatações necessarias
    if origem == 'ae':
        new_name_col = {
            'Proc.RFB-Cont':'Certificado', 'PostoFiscalContador': 'Contador',
            'PostoFiscalSenha': 'Senha', 'PostoFiscalUsuario': 'Usuario'
        }
        df = df.rename(columns = new_name_col)
        df['Contador'] = df['Contador'].fillna('contribuinte')

    #separa
    if separacao == 'separa':
        tabelas = df.groupby('Certificado') if modelo in _REL_PROC else df.groupby('Contador')
    elif separacao == 'nao_separa':
        tabelas = [('', df)]

    #seleciona o modelo e gera os dados
    lista_dados = []
    for name, group in tabelas:
        if name.lower() == 'consumidor': continue
        if name: name = ' - ' + name
        nome = 'Dados' + name + '.' + extensao
        lista_dados.append([nome, seleciona_modelo(modelo, group)])

    #gera e salva os arquivo
    caminho = os.path.join(_RELATORIOS, f'Modelo_{modelo}')
    os.makedirs(caminho, exist_ok=True)
    if extensao == 'csv':
        for nome_arq, dados in lista_dados:
            if not dados: continue
            if not salvar_csv(os.path.join(caminho, nome_arq), dados, modelo, saida): return False
    else:
        for nome_arq, dados in lista_dados:
            if not dados: continue
            if not salvar_py(os.path.join(caminho, nome_arq), dados, modelo): return False

    print("Tempo de execução: ", datetime.now() - comeco)
    return True


@eel.expose
def busca():
    global _RELATORIOS

    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder = askdirectory(title="Selecione a pasta")
    if not folder: 
        return _RELATORIOS if _RELATORIOS != '.' else "Selecione o caminho para salvar arquivo"
    _RELATORIOS = folder

    lst = [x for x in folder.split('/') if x]
    if len(lst) > 4:
        folder = f'{lst[0]}/{lst[1]}/.../{lst[-2]}/{lst[-1]}'

    return folder


@eel.expose
def get_file():
    global _COMPLEMENTAR

    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    file = askopenfilename(title="Selecione o arquivo", filetypes=[("Excel files","*.xlsx *.xls")])
    if not file: 
        return _COMPLEMENTAR if _COMPLEMENTAR != '.' else "Selecione o arquivo complementar"
    _COMPLEMENTAR = file

    lst = [x for x in file.split('/') if x]
    if len(lst) > 4:
        file = f'{lst[0]}/{lst[1]}/.../{lst[-2]}/{lst[-1]}'

    return file


def handle_exit(ar1, ar2):
    import sys
    print('>>>Finalizado')
    sys.exit(0)

eel.start('index.html', port=0, close_callback=handle_exit)