from win32com import client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
import os

s = Service('V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe')

_msgs = {
    "not_download": "Não foi possivel baixar o chromeDriver correpondente ao chromeBrowser",
    "not_close": "Verifique se existe algum arquivo chromedriver.exe na pasta" + os.getcwd(),
    "not_found": "Chrome browser não esta na pasta padrão C:\\Program Files\\Google\\Chrome\\Application",
}


def get_chrome_version():
    paths = (
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    )

    parser = client.Dispatch("Scripting.FileSystemObject")
    for file in paths:
        try:
            return parser.GetFileVersion(file)
        except Exception:
            continue
    return None


def download_chrome(version):
    import requests, wget, zipfile
    version = '.'.join(version.split('.')[:-1])

    try:
        os.remove('chromedriver.zip')
    except OSError:
        pass

    url = f'https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{version}'
    response = requests.get(url)
    version = response.text

    download_url = f"https://chromedriver.storage.googleapis.com/{version}/chromedriver_win32.zip"
    latest_driver_zip = wget.download(download_url, 'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.zip')

    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall('V:\Setor Robô\Scripts Python\_comum\Chrome driver')

    os.remove(latest_driver_zip)


def initialize_chrome(options=None):
    print('>>> Inicializando Chromedriver...')

    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    if "chromedriver.exe" in r'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe':
        return True, webdriver.Chrome(service=s, options=options)

    version = get_chrome_version()  # get chromeBrowser version
    if version:
        try:
            download_chrome(version)  # download chromeDriver version that matches
        except PermissionError as e:
            print(str(e))
            return False, _msgs["not_close"]
        except Exception as e:
            print(str(e))
            return False, _msgs["not_download"]
    else:
        return False, _msgs["not_found"]

    return True, webdriver.Chrome(service=s, options=options)
_initialize_chrome = initialize_chrome


def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass
_send_input = send_input


def send_select(elem_id, data, driver):
    '''try:'''
    select = Select(driver.find_element(by=By.ID, value=elem_id))
    select.select_by_value(data)
    '''except:
        pass'''
_send_select = send_select
