# -*- coding: utf-8 -*-
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium import webdriver

from warnings import filterwarnings
from bs4 import BeautifulSoup
from pathlib import Path
from time import sleep
from PIL import Image
import os, shutil, uuid

# variaveis globais
_base = Path('.')

# mensagens de erro curtas, relevantes e genericas
# para serem reutilizaveis
_errors = {
    "ndownload": "Não baixou .exe do driver",
    "nfound": "Chrome não estava em C:\\Program Files",
    "ntype": "Tipo de driver informado não foi encontrado",
    "nclose": f"Existe um .exe do driver na pasta {os.getcwd()} ?",
}


# Procura indefinidamente por elemento com id 'elem_id'
# e envia os dados 'data' quando encontra
# Retorna None
def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass
_send_input = send_input


# Procura por elemento select com find_by[0], sendo find_by[0] id ou name, igual a find_by[1]
# e seleciona a opcao com slt_by[0], sendo slt_by[0] value ou text, igual a slt_by[1]
# Retorna True caso sucesso, retorna False caso não encontre o elemento
# Estoura uma excecao personalizada caso find_by não seja 'id' ou 'name'
def send_select(find_by, slt_by, driver, delay=0):
    try:
        if find_by[0].lower() == 'id':
            slt = Select(driver.find_element(by=By.ID, value=find_by[1]))
        elif find_by[0].lower() == 'name':
            slt = Select(driver.find_element(by=By.ID, value=find_by[1]))
        else:
            raise Exception(f'erro: find_by {find_by[0]}, deve ser id ou name')
    except NoSuchElementException:     
        return False

    if slt_by[0].lower() == 'value':
        slt.select_by_value(slt_by[1])
    elif slt_by[0].lower() == 'text':
        slt.select_by_visible_text(slt_by[1])
    else:
        raise Exception(f'erro: slt_by "{slt_by[0]}", deve ser id ou name')

    sleep(delay)
    return True
_send_select = send_select


# Obtem o dicionario retangulo do elemento 'elem' obtem um
# corte da image que corresponda as coordenadas do 'elem'
# Retorna o nome do arquivo png cropado caso sucesso
# Levanta uma excecao caso erro
'''def get_elem_img(elem, driver, img_path=r'ignore\temp'):
    os.makedirs(img_path, exist_ok=True)
    name = os.path.join(img_path, str(uuid.uuid4().hex) + '.png')

    x, y, width, height = elem.rect.values()
    x, y, width, height = width, height, width + y, height + x

    driver.save_screenshot(name)
    with Image.open(name) as img:
        aux = img.crop((x, y, width, height))
        aux.save(name)

    return name
_get_elem_img = get_elem_img'''


# Obtem o caminho do executal do webdriver, abstraindo
# o tipo de webdriver ou o sistema operacional
# Retorna string caminho do driver em caso de sucesso
# Retorna None em caso de erro
def check_driver_exe(tipo, version=None):
    driver_path = _base / 'ignore' / 'driver'
    os.makedirs(driver_path, exist_ok=True)

    if version:
        version = '.'.join(version.split('.')[:-1])

    for arq in driver_path.iterdir():
        if tipo not in str(arq): continue
        if version and version not in str(arq): continue

        cond = (
            os.name != 'posix' and 'win32' in str(arq),
            os.name == 'posix' and 'linux64' in str(arq)
        )

        if not any(cond): continue
        exe = str(arq)
        break
    else:
        return None

    try:
        # concede permissão de read, write e execute
        # para o proprietario, o grupo e outros
        os.chmod(exe, 0o777)
    except OSError:
        pass

    return exe
_check_driver_exe = check_driver_exe


# Retorna versão chromeBrowser ou None
def version_chrome():
    from subprocess import check_output

    if os.name == 'posix':
        try:
            version = check_output(["google-chrome", "--version"])
            return version.decode().split()[-1]
        except FileNotFoundError:
            pass
    else:
        try:
            reg = r"HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon"
            version = check_output(["reg", "query", f"{reg}", "/v", "version"])
            return version.decode().split()[-1]
        except:
            pass

    return None
_version_chrome = version_chrome


# download chromedriver que correspondente a versão
# Retorna None
def download_chrome(version):
    import requests, wget, zipfile
    print(">>> download chromedriver")

    if os.name == 'posix':
        tipo, exe = ('linux64', 'chromedriver')
    else:
        tipo, exe = ('win32', 'chromedriver.exe')

    url_base = 'https://chromedriver.storage.googleapis.com'
    version = '.'.join(version.split('.')[:-1])

    chromezip = 'chromedriver.zip'
    if (_base / chromezip).is_file():
        os.remove(chromezip)

    url = f'{url_base}/LATEST_RELEASE_{version}'
    version = requests.get(url).text

    url = f"{url_base}/{version}/chromedriver_{tipo}.zip"
    wget.download(url, chromezip)

    with zipfile.ZipFile(chromezip, 'r') as file:
        file.extractall()

    os.makedirs(Path('ignore', 'driver'), exist_ok=True)
    shutil.move(str(_base / exe), str(_base / 'ignore' / 'driver' / f'{tipo}-{exe}-{version}'))

    os.remove(chromezip)
_download_chrome = download_chrome


# Cria uma instancia de webdriver utilizando o chromedriver com as
# opções 'opts' mais as opções default definidas aqui, caso opts seja None
# utiliza as opções a opção --start-maximized mais as opções default.
# Retorna True e a instancia webdriver criada caso sucesso
# Retorna False e uma mensagem caso falha
def exec_chrome(user_opts=None):
    print(">>> inicializando chromedriver")

    # opções default
    opts = webdriver.ChromeOptions()
    opts.add_argument('--headless')
    opts.add_argument("--window-size=1600,900")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_experimental_option('excludeSwitches', ['enable-logging'])

    if not user_opts:
        opts.add_argument("--start-maximized")
    else:
        for arg in user_opts.arguments:
            opts.add_argument(arg)

    version = version_chrome()
    if not version: return _errors["nfound"]

    exe = check_driver_exe('chrome', version)
    if exe: return webdriver.Chrome(executable_path=exe, options=opts)

    try:
        download_chrome(version)
    except PermissionError as e:
        return _errors["nclose"]
    except Exception as e:
        print(str(e))
        return _errors["ndownload"]

    exe = check_driver_exe('chrome', version)
    return webdriver.Chrome(executable_path=exe, options=opts)
_exec_chrome = exec_chrome


# Obtem a url de download do pahntomjs e o nome do arquivo
def get_url_phatomjs(content):
    tipo = 'linux-64-bit' if os.name == 'posix' else 'windows'
    soup = BeautifulSoup(content, 'html.parser')

    url = soup.find('h2', attrs={'id': tipo})
    url = url.find_next("a").get('href', '')

    return url, url.split('/')[-1]


# download versão mais recente do phantomjs
def download_phantomjs():
    import requests, wget, tarfile, zipfile
    print(">>> download phantom js")

    res = requests.get('https://phantomjs.org/download.html')
    url, file = get_url_phatomjs(res.content)
    if not all((url, file)): return False

    phantomzip = _base / file 
    if phantomzip.is_file(): os.remove(phantomzip)

    wget.download(url, str(phantomzip))
    
    path_driver = _base / 'ignore' / 'driver'
    os.makedirs(path_driver, exist_ok=True)

    if os.name == 'posix':
        arq = tarfile.open(phantomzip, 'r:bz2')
        driver, aux = 'phantomjs', 'linux64'
    else:
        arq = zipfile.ZipFile(phantomzip, 'r')
        driver, aux = 'phantomjs.exe', 'win32'

    root_file = file.rstrip('.zip').rstrip('.tar.bz2')
    path_file = str(Path(root_file, 'bin', driver))

    # zip files sempre usam slash / como separador
    #
    arq.extract(f'{root_file}/bin/{driver}')
    arq.close()

    shutil.move(path_file, (path_driver / f'{aux}-{driver}'))
    shutil.rmtree(root_file)

    os.remove(phantomzip)
    return True
_download_phantomjs = download_phantomjs


# Cria uma instancia de webdriver utilizando o phantom js com 
# um certificado p12 'cert' caso haja, cabeçalho 'headers' caso haja
# e opções 'args' caso haja mais a opçãos default '--ignore-ssl-errors=true'
# Retorna True e a instancia webdriver criada caso sucesso
# Retorna False e uma mensagem caso falha
def exec_phantomjs(cert=None, user_args=None, headers=None, size=(1600, 900)):
    print(">>> inicializando phantom js")
    filterwarnings('ignore')

    if headers:
        for key, value in enumerate(headers):
            capability_key = 'phantomjs.page.customHeaders.{}'.format(key)
            webdriver.DesiredCapabilities.PHANTOMJS[capability_key] = value

    args = ['--ignore-ssl-errors=true']
    if user_args: args.extend(user_args)

    if cert: args.append(f'--ssl-client-certificate-file={cert}')

    exe = check_driver_exe('phantom')
    if exe:
        driver = webdriver.PhantomJS(executable_path=exe, service_args=args)
        driver.set_window_size(*size)
        driver.delete_all_cookies()
        return driver

    try:
        download_phantomjs()
    except PermissionError as e:
        return _errors["nclose"]
    except Exception as e:
        print(str(e))
        return _errors["ndownload"]

    exe = check_driver_exe('phantom')
    if not exe: return _errors['ntype']

    driver = webdriver.PhantomJS(executable_path=exe, service_args=args)
    driver.set_window_size(*size)
    driver.delete_all_cookies()
    return driver
_exec_phantomjs = exec_phantomjs