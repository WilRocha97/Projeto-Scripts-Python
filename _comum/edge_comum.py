from win32com import client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.select import Select
import os

s = Service('V:\Setor Robô\Scripts Python\_comum\Edge driver\msedgedriver.exe')

_msgs = {
    "not_download": "Não foi possivel baixar o edgeDriver correspondente ao edgeBrowser",
    "not_close": "Verifique se existe algum arquivo msedgedriver.exe na pasta" + os.getcwd(),
    "not_found": "Edge browser não esta na pasta padrão C:\\Program Files\\Google\\Edge\\Application",
}


def initialize_edge(options=None):
    print('>>> Inicializando Edgedriver...')

    if not options:
        options = webdriver.EdgeOptions()
        options.use_chromium = True
        options.add_argument("--start-maximized")

    options.add_argument("--ignore-certificate-errors")

    return True, webdriver.Edge(service=s, options=options)
_initialize_edge = initialize_edge


def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.ID, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass
_send_input = send_input


def send_input_xpath(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_id)
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


def find_by_class(iten, driver):
    try:
        elem = driver.find_element(by=By.CLASS_NAME, value=iten)
        return elem
    except:
        return None
_find_by_class = find_by_class
