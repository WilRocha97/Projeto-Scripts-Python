# -*- coding: utf-8 -*-
from selenium import webdriver

from warnings import filterwarnings
from bs4 import BeautifulSoup
from pathlib import Path
from time import sleep
from PIL import Image
import os, shutil, uuid

# variáveis globais
_base = Path('.')

# mensagens de erro curtas, relevantes e genéricas
# para serem reutilizáveis
_errors = {
    "ndownload": "Não baixou .exe do driver",
    "nfound": "Chrome não estava em C:\\Program Files",
    "ntype": "Tipo de driver informado não foi encontrado",
    "nclose": f"Existe um .exe do driver na pasta {os.getcwd()} ?",
}


# Obtém a url de download do pahntomjs e o nome do arquivo
def get_url_phantomjs(content):
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


# Cria uma instância de webdriver utilizando o phantom js com
# um certificado p12 'cert' caso haja, cabeçalho 'headers' caso haja
# e opções 'args' caso haja mais a opções default '--ignore-ssl-errors=true'
# Retorna True e a instância webdriver criada caso sucesso
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
    except PermissionError:
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
