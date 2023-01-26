"""from 'twocaptcha.solver' import TimeoutException, ValidationException
from twocaptcha.api import ApiException, NetworkException
from twocaptcha import TwoCaptcha"""
from anticaptchaofficial.imagecaptcha import *
from anticaptchaofficial.recaptchav2proxyless import *
from anticaptchaofficial.hcaptchaproxyless import *
from time import sleep
import os

dados = "V:\\Setor Robô\\Scripts Python\\_comum\\Dados.txt"
f = open(dados, 'r', encoding='utf-8')
key = f.read()

# variaveis globais
api_key = os.getenv('APIKEY_2CAPTCHA', '')
anticaptcha_api_key = key


# ---------- VVVV 2 Captcha (Não utilizado) VVVV ---------------------------------------------------------------------------------------------------------------------------


'''def break_recaptcha_v2(data, limit=5):
    tipo = 'recaptcha-v2'
    print(f'>>> quebrando {tipo}')

    for i in range(limit):
        try:
            solver = TwoCaptcha(api_key)
            result = solver.recaptcha(
                sitekey=data['sitekey'], url=data['url'], version='v2'
            )
            return result

        except (TimeoutException, ApiException) as e:
            print(f'\t{str(e)}')

        except (NetworkException, ValidationException) as e:
            return 'Erro captcha - ' + str(e)
    else:
        return f'Erro captcha - não resolveu {tipo}'
_break_recaptcha_v2 = break_recaptcha_v2


# Envia os dados do recaptcha para a api two capctha
# Retorna um dicionario com id da requisição e token do capctha
# Retorna uma string/mensagem em caso de erro
def break_recaptcha_v3(data, limit=5):
    tipo = 'recaptcha-v3'
    print(f'>>> quebrando {tipo}')

    for i in range(limit):
        try:
            solver = TwoCaptcha(api_key)
            result = solver.recaptcha(
                sitekey=data['sitekey'], url=data['url'],
                action=data['action'], version='v3'
            )
            return result

        except (TimeoutException, ApiException) as e:
            print(f'\t{str(e)}')

        except (NetworkException, ValidationException) as e:
            return 'Erro captcha - ' + str(e)
    else:
        return f'Erro captcha - não resolveu {tipo}'
_break_recaptcha_v3 = break_recaptcha_v3


# Envia os dados do hcaptcha para a api two capctcha
# Retorna um dicionario com id da requisição e token do capctha
# Retorna uma string/mensagem em caso de erro
def break_hcaptcha(data, limit=5):
    tipo = 'hcaptcha'
    print(f'>>> quebrando {tipo}')

    for i in range(limit):
        try:
            solver = TwoCaptcha(api_key)
            result = solver.hcaptcha(
                sitekey=data['sitekey'], url=data['url']
            )
            return result

        except (TimeoutException, ApiException) as e:
            print(f'\t{str(e)}')

        except (NetworkException, ValidationException) as e:
            return 'Erro captcha - ' + str(e)
    else:
        return f'Erro captcha - não resolveu {tipo}'
_break_hcaptcha = break_hcaptcha


# Envia imagem do normal captcha para a api two capctcha
# Retorna um dicionario com id da requisição e token do capctha
# Retorna uma string/mensagem em caso de erro
def break_normal_captcha(img, limit=5):
    tipo = 'normal-captcha'
    print(f'>>> quebrando {tipo}')

    for i in range(limit):
        try:
            solver = TwoCaptcha(api_key)
            result = solver.normal(img)
            return result

        except (TimeoutException, ApiException) as e:
            print(f'\t{str(e)}')

        except (NetworkException, ValidationException) as e:
            return 'Erro captcha - ' + str(e)
    else:
        return f'Erro captcha - não resolveu {tipo}'
_break_normal_captcha = break_normal_captcha'''


# ---------- VVVV Anti Captcha VVVV ---------------------------------------------------------------------------------------------------------------------------


# Recebe a url do site com o captcha mais a chave da api resposnável pelo captcha encontrada no código do site
# envia para a api e retorna a chave para resolver o captcha
def solve_recaptcha(data):
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
_solve_recaptcha = solve_recaptcha


# Recebe a imagem do captcha e envia para a api e retorna o texto da imagem
def solve_text_captcha(img):
    print('>>> Quebrando text captcha')
    solver = imagecaptcha()
    solver.set_verbose(1)
    solver.set_key(anticaptcha_api_key)

    captcha_text = solver.solve_and_return_solution(img)
    if captcha_text != 0:
        return captcha_text
    else:
        print(solver.error_code)
        return solver.error_code
_solve_text_captcha = solve_text_captcha


# Recebe a url do site com o captcha mais a chave da api resposnável pelo captcha encontrada no código do site
# envia para a api e retorna a chave para resolver o captcha
def solve_hcaptcha(data):
    response = 'ERROR_NO_SLOT_AVAILABLE'
    while response == 'ERROR_NO_SLOT_AVAILABLE':
        print('>>> Quebrando hcaptcha')
        solver = hCaptchaProxyless()
        solver.set_verbose(1)
        solver.set_key(anticaptcha_api_key)
        solver.set_website_url(data['url'])
        solver.set_website_key(data['sitekey'])
    
        # tell API that Hcaptcha is invisible
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
_solve_hcaptcha = solve_hcaptcha
