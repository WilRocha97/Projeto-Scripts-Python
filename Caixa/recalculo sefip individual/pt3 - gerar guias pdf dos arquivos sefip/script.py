from time import sleep
import zipfile as zipf
import pyautogui as a
import shutil, os, sys

sys.path.append(r'..\..\..\_comum')
from pyautogui_comum import wait_img, click_img, find_img, focus


# variaveis globais
_path_sfp = os.path.join('c:\\','program files (x86)','caixa','sefip')

_base = os.path.join(r'\\vpsrv02','dca','sefip','arquivos')
_path_end = os.path.join(_base, 'guias')
_path_pdf = os.path.join(_base, 'pdf')


def compactar_guias():
    zip_name = os.path.join(_path_end, 'guias_SEFIP.zip')

    with zipf.ZipFile(zip_name, 'w', zipf.ZIP_DEFLATED) as z:
        for arq in os.listdir(_path_end):
            if arq[-4:].lower() != '.pdf': continue

            z.write(os.path.join(_path_end, arq))
            print(f">>>{arq} salvo")


def emitir_guia(xml):
    click_img('search_sefip.png', conf=0.9)
    sleep(1)

    a.press(('alt', 'alt','r','g','i'), interval=0.5)
    img = 'abrir_icp.png'
    if not wait_img(img, conf=0.9):
        return False, f'>>>{xml} Não achou imagem {img}'

    a.write(os.path.join(_path_pdf, xml))
    a.press('enter')
    while not find_img('gerar_pdf.png', conf=0.9):
        sleep(1)
        if find_img('erro_gerar.png', conf=0.9):
            sleep(2)
            a.press('enter')
            registra_ocorrencia(xml)
            return True, f'{xml} Não existe GRF para impressão'

    a.hotkey('alt','p')
    img = 'salvar_pdf.png'
    if not wait_img(img, conf=0.9):
        return False, f'>>>{xml} Não achou imagem {img}'

    a.press('enter')
    img = 'info_sucesso.png'
    if not wait_img(img, conf=0.9):
        return False, f'>>>{xml} Não achou imagem {img}'

    a.press('enter')
    a.hotkey('alt','f')
    a.press('alt')
    return True, f'>>>{xml} guia de gerada'

    
def run():
    os.makedirs(_path_end, exist_ok=True)
    
    res, msg = focus(title='SEFIP -', maximize=False)
    if not res: raise Exception(msg)

    print('------Emitir Guias------')
    for arq in os.listdir(_path_pdf):
        if arq[-4:].lower() != '.xml':
            continue

        status, msg = emitir_guia(arq)
        print(msg)

    print('-------Mover Guias-------')
    for arq in os.listdir(_path_sfp):
        if arq[-4:].lower() != '.pdf': continue

        aux = os.path.join(_path_sfp, arq)
        shutil.copy(aux, _path_end)
        # os.remove(aux)
        print(f'>>>{arq} movido')

    print('-------Mover Comprovantes-------')
    for arq in os.listdir(_path_pdf):
        if arq[-4:].lower() != '.pdf': continue

        aux = os.path.join(_path_pdf, arq)
        shutil.copy(aux, _path_end)
        os.remove(aux)
        print(f'>>>{arq} movido')

    print('-----Compactar Guias-----')
    compactar_guias()


if __name__ == '__main__':
    run()