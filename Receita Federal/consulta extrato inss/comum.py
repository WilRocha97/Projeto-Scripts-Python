from win32com.client import Dispatch
from selenium import webdriver
import os

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

    parser = Dispatch("Scripting.FileSystemObject")
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
    latest_driver_zip = wget.download(download_url,'chromedriver.zip')

    with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
        zip_ref.extractall()

    os.remove(latest_driver_zip)


def initialize_chrome(options=None):
    print('>>> Inicializando chromedriver...')

    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    if "chromedriver.exe" in os.listdir(os.getcwd()):
        return True, webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)

    version = get_chrome_version() # get chromeBrowser version
    if version:
        try:
            download_chrome(version) # download chromeDriver version that matches
        except PermissionError as e:
            print(str(e))
            return False, _msgs["not_close"]
        except Exception as e:
            print(str(e))
            return False, _msgs["not_download"]
    else:
        return False, _msgs["not_found"]

    return True, webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)