"""from anticaptchaofficial.imagecaptcha import *
from anticaptchaofficial.recaptchav2proxyless import *
from anticaptchaofficial.recaptchav3proxyless import *
from anticaptchaofficial.hcaptchaproxyless import *"""
from selenium import webdriver
from twocaptcha import TwoCaptcha
from selenium.webdriver.common.by import By
from PIL import Image
from time import sleep
import os, sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))


dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados Anti.txt"
f = open(dados, 'r', encoding='utf-8')
key_anti = f.read()
anticaptcha_api_key = key_anti

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados 2Cap.txt"
f = open(dados, 'r', encoding='utf-8')
key_2cap = f.read()
api_key = os.getenv('APIKEY_2CAPTCHA', key_2cap)



# ---------- VVVV 2 Captcha VVVV ---------------------------------------------------------------------------------------------------------------------------


def solve_recaptcha(data):
    tipo = 'recaptcha-v2'
    print(f'>>> Resolvendo {tipo}')

    solver = TwoCaptcha(api_key)
    result = solver.recaptcha(sitekey=data['sitekey'], url=data['url'])
    
    print('>>> Captcha resolvido')
    return str(result['code'])
_solve_recaptcha = solve_recaptcha


# Envia os dados do recaptcha para a api two captcha
# Retorna um dicionario com id da requisição e token do captcha
# Retorna uma string/mensagem em caso de erro
def solve_recaptcha_v3(data):
    tipo = 'recaptcha-v3'
    print(f'>>> Resolvendo {tipo}')

    solver = TwoCaptcha(api_key)
    result = solver.recaptcha(sitekey=data['sitekey'], url=data['url'], version='v3', action='test')
    
    print('>>> Captcha resolvido')
    return str(result['code'])
_solve_recaptcha_v3 = solve_recaptcha_v3


# Envia os dados do hcaptcha para a api two capctcha
# Retorna um dicionário com id da requisição e token do captcha
# Retorna uma string/mensagem em caso de erro
def solve_hcaptcha(data):
    tipo = 'hcaptcha'
    print(f'>>> Resolvendo {tipo}')
    
    solver = TwoCaptcha(api_key)
    result = solver.hcaptcha(sitekey=data['sitekey'], url=data['url'])
    
    print('>>> Captcha resolvido')
    return str(result['token'])
_solve_hcaptcha = solve_hcaptcha


# Envia imagem do normal captcha para a api two captcha
# Retorna um dicionário com id da requisição e token do captcha
# Retorna uma string/mensagem em caso de erro
def solve_text_captcha(driver, captcha_element, element_type='id'):
    os.makedirs('ignore\captcha', exist_ok=True)
    # captura a imagem do captcha
    if element_type == 'id':
        element = driver.find_element(by=By.ID, value=captcha_element)
    elif element_type == 'xpath':
        element = driver.find_element(by=By.XPATH, value=captcha_element)
    
    location = element.location
    size = element.size
    driver.save_screenshot('ignore\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    time.sleep(2)
    im = Image.open(r'ignore\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'ignore\captcha\captcha.png')
    time.sleep(1)
    
    print('>>> Resolvendo text captcha')

    solver = TwoCaptcha(api_key)
    result = solver.normal(os.path.join('ignore', 'captcha', 'captcha.png'))
    
    print('>>> Captcha resolvido')
    return str(result['text'])
_solve_text_captcha = solve_text_captcha


def solve_audio_recaptcha(driver, iframe, id_input_captcha, id_confirma_captcha):
    os.makedirs('ignore\captcha', exist_ok=True)
    caminho_audio = 'ignore\captcha\\audio.mp3'
    # IMPORTAÇÃO BIBLIOTECA QUE TRANSFORMA AUDIO EM TEXTO
    print('>>> Resolvendo recaptcha audio')

    #_click_img('audio.png', conf=0.9)
    #time.sleep(2)
    #_click_img('download_audio.png', conf=0.9)
    #time.sleep(2)
    #_click_img('3pontos.png', conf=0.9)
    #time.sleep(2)
    #_click_img('download2.png', conf=0.9)
    #time.sleep(2)
    #p.hotkey('ctrl', 'w')
    #time.sleep(2)

    solver = TwoCaptcha(api_key)
    result = solver.audio(caminho_audio, lang='pt')

    # TROCA O FRAME PARA LOCALIZAR O CAMPO DE RESPOSTA DO CAPTCHA
    driver.switch_to.frame(driver.find_element(by=By.XPATH, value=iframe))
    time.sleep(1)

    # DIGITA NO CAMPO DE RESPOSTA DO CAPTCHA DO AUDIO
    driver.find_element(by=By.ID, value=id_input_captcha).send_keys(result['text'])

    # CLICKA NO BOTÂO VERIFICAR DO CAPTCHA
    driver.find_element(by=By.ID, value=id_confirma_captcha).click()

    # EXCLUI O ARQUIVO DE AUDIO BAIXADO
    os.remove(caminho_audio)

    #if _find_img('refaz.png', conf=0.9):
        #REFAZ O CAPTCHA DENOVO PQ DEU ERRADO!

    # VOLTA O FRAME AO ORIGINAL
    driver.switch_to.default_content()
_solve_audio_recaptcha = solve_audio_recaptcha


# ---------- VVVV Anti Captcha VVVV ---------------------------------------------------------------------------------------------------------------------------


# Recebe a url do site com o captcha mais a chave da api responsável pelo captcha encontrada no código do site
# envia para a api e retorna a chave para resolver o captcha
"""def solve_recaptcha(data):
    print('>>> Quebrando recaptcha')
    solver = recaptchaV2Proxyless()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    solver.set_website_url(data['url'])
    solver.set_website_key(data['sitekey'])
    # set optional custom parameter which Google made for their search page Recaptcha v2
    # solver.set_data_s('"data-s" token from Google Search results "protection"')

    g_response = solver.solve_and_return_solution()
    if g_response != 0:
        return g_response
    else:
        print(solver.error_code)
        return solver.error_code
_solve_recaptcha = solve_recaptcha"""


"""def solve_recaptcha_v3(data):
    print('>>> Quebrando recaptcha v3')
    solver = recaptchaV3Proxyless()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)
    solver.set_website_url(data['url'])
    solver.set_website_key(data['sitekey'])
    solver.set_page_action("home_page")
    solver.set_min_score(0.9)
    
    # Specify softId to earn 10% commission with your app.
    # Get your softId here: https://anti-captcha.com/clients/tools/devcenter
    solver.set_soft_id(0)
    
    g_response = solver.solve_and_return_solution()
    if g_response != 0:
        return g_response
    else:
        print(solver.error_code)
        return solver.error_code
_solve_recaptcha_v3 = solve_recaptcha_v3"""


# Recebe a imagem do captcha e envia para a api e retorna o texto da imagem
"""def solve_text_captcha(driver, captcha_element, element_type='id'):
    os.makedirs('ignore\captcha', exist_ok=True)
    # captura a imagem do captcha
    if element_type == 'id':
        element = driver.find_element(by=By.ID, value=captcha_element)
    elif element_type == 'xpath':
        element = driver.find_element(by=By.XPATH, value=captcha_element)
    
    location = element.location
    size = element.size
    driver.save_screenshot('ignore\captcha\pagina.png')
    x = location['x']
    y = location['y']
    w = size['width']
    h = size['height']
    width = x + w
    height = y + h
    time.sleep(2)
    im = Image.open(r'ignore\captcha\pagina.png')
    im = im.crop((int(x), int(y), int(width), int(height)))
    im.save(r'ignore\captcha\captcha.png')
    time.sleep(1)
    
    print('>>> Quebrando text captcha')
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)

    captcha_text = solver.solve_and_return_solution(os.path.join('ignore', 'captcha', 'captcha.png'))
    if captcha_text != 0:
        return captcha_text
    else:
        print(solver.error_code)
        return solver.error_code
_solve_text_captcha = solve_text_captcha"""


# Recebe a url do site com o captcha mais a chave da api responsável pelo captcha encontrada no código do site
# envia para a api e retorna a chave para resolver o captcha
"""def solve_hcaptcha(data, visible=False):
    response = 'ERROR_NO_SLOT_AVAILABLE'
    while response == 'ERROR_NO_SLOT_AVAILABLE':
        print('>>> Quebrando hcaptcha')
        solver = hCaptchaProxyless()
        solver.set_verbose(1)
        solver.set_key(anticaptcha_api_key)
        solver.set_website_url(data['url'])
        solver.set_website_key(data['sitekey'])
    
        # tell API that Hcaptcha is invisible
        if not visible:
            solver.set_is_invisible(1)
    
        # set here parameters like rqdata, sentry, apiEndpoint, endpoint, reportapi, assethost, imghost
        # solver.set_enterprise_payload({
        #    "rqdata": "rq data value from target website",
        #    "sentry": True
        # })
    
        g_response = solver.solve_and_return_solution()
        
        if g_response != 0:
            return g_response
        else:
            if solver.error_code == 'ERROR_NO_SLOT_AVAILABLE':
                sleep(5)
                response = 'ERROR_NO_SLOT_AVAILABLE'
            else:
                print(solver.error_code)
                return solver.error_code
_solve_hcaptcha = solve_hcaptcha"""


"""def solve_audio_recaptcha(driver, iframe, id_input_captcha, id_confirma_captcha):
    os.makedirs('ignore\captcha', exist_ok=True)
    caminho_audio = 'ignore\captcha\\audio.mp3'
    # IMPORTAÇÃO BIBLIOTECA QUE TRANSFORMA AUDIO EM TEXTO
    print('>>> Quebrando recaptcha audio')
    import whisper

    #_click_img('audio.png', conf=0.9)
    #time.sleep(2)
    #_click_img('download_audio.png', conf=0.9)
    #time.sleep(2)
    #_click_img('3pontos.png', conf=0.9)
    #time.sleep(2)
    #_click_img('download2.png', conf=0.9)
    #time.sleep(2)
    #p.hotkey('ctrl', 'w')
    #time.sleep(2)

    model = whisper.load_model("base")
    result = model.transcribe(caminho_audio)

    # TROCA O FRAME PARA LOCALIZAR O CAMPO DE RESPOSTA DO CAPTCHA
    driver.switch_to.frame(driver.find_element(by=By.XPATH, value=iframe))
    time.sleep(1)

    # DIGITA NO CAMPO DE RESPOSTA DO CAPTCHA DO AUDIO
    driver.find_element(by=By.ID, value=id_input_captcha).send_keys(result['text'])

    # CLICKA NO BOTÂO VERIFICAR DO CAPTCHA
    driver.find_element(by=By.ID, value=id_confirma_captcha).click()

    # EXCLUI O ARQUIVO DE AUDIO BAIXADO
    os.remove(caminho_audio)

    #if _find_img('refaz.png', conf=0.9):
        #REFAZ O CAPTCHA DENOVO PQ DEU ERRADO!

    # VOLTA O FRAME AO ORIGINAL
    driver.switch_to.default_content()

_solve_audio_recaptcha = solve_audio_recaptcha"""