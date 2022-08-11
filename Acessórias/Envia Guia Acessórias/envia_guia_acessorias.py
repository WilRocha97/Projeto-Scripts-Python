from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from sys import path
path.append(r'..\..\_comum')
from comum_comum import _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _indice
from chrome_comum import initialize_chrome, _send_input


def find_by_id(elemento, driver):
    try:
        elem = driver.find_element(by=By.ID, value=elemento)
        return elem
    except:
        return None


def login(driver):
    print('>>> Abrindo o site')
    url = 'https://app.acessorias.com/?login'
    driver.get(url)
    time.sleep(1)
    
    driver.find_element(by=By.NAME, value='mailAC').send_keys('joao@veigaepostal.com.br')
    driver.find_element(by=By.NAME, value='passAC').send_keys('Milenio03')
    time.sleep(1)
    driver.find_element(by=By.XPATH, value='/html/body/main/section[1]/div/form/div[2]/button').click()
    
    while not find_by_id('M144', driver):
        time.sleep(1)

    return driver


def monitorar(driver):
    url_departamento = [('CONTÁBIL', 'https://app.acessorias.com/sysmain.php?m=144&act=e&i=1'), ('IRRF', 'https://app.acessorias.com/sysmain.php?m=144&act=e&i=26'),
                        ('DEPARTAMENTO PESSOAL', 'https://app.acessorias.com/sysmain.php?m=144&act=e&i=17'), ('FISCAL', 'https://app.acessorias.com/sysmain.php?m=144&act=e&i=9')]
    
    for departamento in url_departamento:
        driver.get(departamento[1])
        # print(f'>>> Acessando o departamento {departamento[0]}')
        while not find_by_id('DptNome', driver):
            time.sleep(1)
        
        if 'com guias agrupadas para envio nesse departamento.' in driver.page_source:
            driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[2]/div[2]/div/div/label/a').click()
            # print(f'>>> Abrindo guias')
            while 'Enviar agora' not in driver.page_source:
                time.sleep(1)

            # print(f'>>> Enviando guias')
            driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[2]/div[2]/div/div/div[3]/div[1]/span/a').click()
            while 'Confirma enviar o e-mail agendado agora?' not in driver.page_source:
                time.sleep(1)

            driver.find_element(by=By.XPATH, value='/html/body/div[6]/div/div[3]/button[1]').click()
            while 'OK, envio realizado com sucesso!' not in driver.page_source:
                time.sleep(0.5)
        

@_time_execution
def run():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--window-size=1920,1080')
    
    status, driver = initialize_chrome()
    
    driver_logado = login(driver)
    while 1 > 0:
        monitorar(driver_logado)


if __name__ == '__main__':
    run()
