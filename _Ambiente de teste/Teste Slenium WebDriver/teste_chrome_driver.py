# Ambiente para testar se um site funciona no chrome driver

import time
from selenium import webdriver
from selenium.webdriver.common.by import By

from sys import path
path.append(r'..\..\_comum')
from chrome_comum import _initialize_chrome

options = webdriver.ChromeOptions()
# options.add_argument('--headless')
# options.add_argument('--window-size=1920,1080')
options.add_argument("--start-maximized")


status, driver = _initialize_chrome(options)
driver.get('https://www.google.com')
while True:
    time.sleep(1)
