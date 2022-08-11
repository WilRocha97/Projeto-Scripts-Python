# -*- coding: utf-8 -*-
from urllib3 import disable_warnings, exceptions
from threading import Thread
from time import sleep
import requests
import eel

disable_warnings(exceptions.InsecureRequestWarning)

_ATIVO = True


@eel.expose
def handle_exit(ar1, ar2):
    import sys
    global _ATIVO
    _ATIVO = False

    sys.exit(0)


def get_status(url):
    while _ATIVO:
        try:
            pagina = requests.get(url, verify=False)
            if pagina.status_code == 200:
                eel.status(True)
            else:
                eel.status(False)
        except:
            eel.status(False)

        sleep(1)


@eel.expose
def verifica(url):
    global _ATIVO
    _ATIVO = False
    sleep(2)
    _ATIVO = True
    t = Thread(target=get_status, args=(url, ))
    t.daemon = True
    t.start()

    return True


def main():
    eel.init('web', allowed_extensions=['.css', '.html', '.js'])
    eel.start('index.html', close_callback=handle_exit,  port=0, size=(600,200))


if __name__ == '__main__':
    main()
    