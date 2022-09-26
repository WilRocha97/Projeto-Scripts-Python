from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from time import sleep
import pyautogui as p
import os
import cv2


s = Service('chromedriver.exe')


def initialize_chrome(options=None):
    print('>>> Inicializando Chromedriver...')
    
    if not options:
        options = webdriver.ChromeOptions()
        options.add_argument("start-maximized")

    options.add_argument("--ignore-certificate-errors")
    options.add_experimental_option("prefs", {"profile.default_content_setting_values.notifications": 1})
    options.add_experimental_option('excludeSwitches', ['enable-logging'], )
    
    # if "chromedriver.exe" in os.listdir(os.getcwd()):
    return True, webdriver.Chrome(service=s, options=options)


def send_input(elem_id, data, driver):
    while True:
        try:
            elem = driver.find_element(by=By.XPATH, value=elem_id)
            elem.send_keys(data)
            break
        except:
            pass
    

def login():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = initialize_chrome()

    driver.get('https://www.dominioweb.com.br/')
    send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[1]/span[2]/input', 'robo@veigaepostal.com.br', driver)
    send_input('/html/body/app-root/app-login/div/div/fieldset/div/div/section/form/label[2]/span[2]/input', 'Rb#0086*', driver)
    driver.find_element(by=By.ID, value='enterButton').click()
    
    caminho = os.path.join('imgs', 'abrir_app.png')
    caminho2 = os.path.join('imgs', 'abrir_app_2.png')
    while not p.locateOnScreen(caminho, confidence=0.9):
        if p.locateOnScreen(caminho2, confidence=0.9):
            sleep(1)
            while p.locateOnScreen(caminho2, confidence=0.9):
                p.click(p.locateCenterOnScreen(caminho2, confidence=0.9), button='left')
                return driver

        sleep(1)
    if p.locateOnScreen(caminho, confidence=0.9):
        p.click(p.locateCenterOnScreen(caminho, confidence=0.9), button='left')
        return driver


def abrir_acesso():
    while not p.locateOnScreen(r'imgs/modulos.png', confidence=0.9):
        sleep(1)
        try:
            p.getWindowsWithTitle('Lista de Programas')[0].activate()
        except:
            pass
    sleep(3)
    driver.quit()
    p.click(p.locateCenterOnScreen(r'imgs/escrita_fiscal.png', confidence=0.9), button='left', clicks=2)
    while not p.locateOnScreen(r'imgs/login_modulo.png', confidence=0.9):
        sleep(1)

    p.moveTo(p.locateCenterOnScreen(r'imgs/insere_usuario.png', confidence=0.9))
    local_mouse = p.position()
    p.click(int(local_mouse[0] + 120), local_mouse[1], clicks=2)
    
    sleep(0.5)
    p.press('del', presses=10)
    p.write('robo')
    sleep(0.5)
    p.press('tab')
    sleep(0.5)
    p.press('del', presses=10)
    p.write('Rb#0086*')
    sleep(0.5)
    p.hotkey('alt', 'o')
    while not p.locateOnScreen(r'imgs/onvio.png', confidence=0.9):
        sleep(1)
    
    
def monitora_rotina():
    while 1 > 0:
        sleep(1)
        
        if p.locateOnScreen(r'imgs/reconectar.png', confidence=0.8):
            p.click(p.locaCenterOnScreen(r'imgs/sim.png', confidence=0.8))
        if p.locateOnScreen(r'imgs/iniciar_tarefa_agora.png', confidence=0.8):
            p.click(p.locaCenterOnScreen(r'imgs/iniciar_tarefa_agora.png', confidence=0.8))
            p.hotkey('alt', 'n')
        if p.locateOnScreen(r'imgs/lancamentos_regerados.png', confidence=0.8):
            p.hotkey('alt', 'y')
        if p.locateOnScreen(r'imgs/erros_avisos.png', confidence=0.8):
            p.hotkey('alt', 'n')
            sleep(2)
            p.press('esc', presses=100)
    
    
if __name__ == '__main__':
    driver = login()
    abrir_acesso()
    monitora_rotina()
