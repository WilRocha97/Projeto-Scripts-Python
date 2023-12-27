import pyperclip
from win32com import client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
from pyautogui_comum import _click_img, _find_img
import os, time, pyautogui as p

service = Service(log_path = 'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe')

_msgs = {
    "not_download": "Não foi possivel baixar o chromeDriver correspondente ao chromeBrowser",
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


def initialize_chrome(options=webdriver.ChromeOptions()):
    print('>>> Inicializando Chromedriver...')

    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")

    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    if "chromedriver.exe" in r'V:\Setor Robô\Scripts Python\_comum\Chrome driver\chromedriver.exe':
        return True, webdriver.Chrome(service=service, options=options)

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

    return True, webdriver.Chrome(service=service, options=options)
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


def send_input_xpath(elem_path, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_path)
            elem.send_keys(data)
            break
        except:
            pass
_send_input_xpath = send_input_xpath


def send_select(elem_id, data, driver):
    '''try:'''
    select = Select(driver.find_element(by=By.ID, value=elem_id))
    select.select_by_value(data)
    '''except:
        pass'''
_send_select = send_select


def find_by_id(item, driver):
    try:
        driver.find_element(by=By.ID, value=item)
        return True
    except:
        return False
_find_by_id = find_by_id


def find_by_path(item, driver):
    try:
        driver.find_element(by=By.XPATH, value=item)
        return True
    except:
        return False
_find_by_path = find_by_path


def find_by_class(item, driver):
    try:
        driver.find_element(by=By.CLASS_NAME, value=item)
        return True
    except:
        return False
_find_by_class = find_by_class


# abre o chrome padrão instalado no PC
def abrir_chrome(url, fechar_janela=True, outra_janela=False, anti_travamento=False):
    def abrir_nova_janela():
        time.sleep(0.5)
        os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")
        if outra_janela:
            while _find_img(outra_janela, conf=0.9):
                time.sleep(1)
        
        while not _find_img('google.png', conf=0.9):
            if anti_travamento:
                if _find_img('chrome_travado.png', conf=0.9):
                    _click_img('chrome_travado.png', conf=0.9)

            time.sleep(1)
            if _find_img('restaurar_pagina.png', conf=0.9):
                _click_img('restaurar_pagina.png', conf=0.9)
                time.sleep(1)
                p.press('esc')
                time.sleep(1)
    
    if fechar_janela:
        p.hotkey('win', 'm')
    
    if fechar_janela:
        if _find_img('chrome_aberto.png', conf=0.99):
            _click_img('chrome_aberto.png', conf=0.99, timeout=1)
        else:
            abrir_nova_janela()
    else:
        abrir_nova_janela()
        
        _click_img('google.png', conf=0.9)
        time.sleep(1)
        p.hotkey('alt', 'space', 'x')
        time.sleep(1)
    
    for i in range(4):
        p.click(1000, 51)
        time.sleep(0.5)
        p.hotkey('ctrl', 'a')
        time.sleep(0.5)
        p.press('delete')
        time.sleep(0.2)
        pyperclip.copy(url)
        time.sleep(0.2)
        p.hotkey('ctrl', 'v')

    time.sleep(1)
    p.press('enter')
_abrir_chrome = abrir_chrome
