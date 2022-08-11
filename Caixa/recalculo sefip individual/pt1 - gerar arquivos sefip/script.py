from datetime import datetime
from time import sleep
from PIL import Image
import pyautogui as p
import pytesseract as ocr
import os, re, shutil, sys

sys.path.append(r'..\..\..\_comum')
from comum_comum import time_execution, escreve_relatorio_csv
from pyautogui_comum import find_img, wait_img, click_img, focus


# variaveis globais
_path_local = os.path.join('C:\\', 'Program Files (x86)', 'caixa', 'arquivos')
_path_index = os.path.join(r'\\vpsrv02', 'dca', 'sefip', 'indices')
_path_sfp = os.path.join(r'\\vpsrv02', 'dca', 'sefip', 'arquivos', 'sfp')
_path_bkp = os.path.join(r'\\vpsrv02', 'dca', 'sefip', 'backup')
_path_dp = os.path.join(r'\\SRV001', 'Sefip_Guias', 'RELATORIOS_SEFIP', 'BACKUP RECOLHIMENTO PARCIAL')

ocr.pytesseract.tesseract_cmd = os.path.join('ignore', 'tesseract_ocr', 'tesseract.exe')


def check_venc(venc):
    try:
        d, m, y = venc.split('/')
        datetime(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return False
    return True


def check_dir(razao, comps):
    print('Copiando diretorios\n')
    os.makedirs(_path_bkp, exist_ok=True)
    regex = re.compile(r"(\d{2}.?\d{4}).*(\.[zipZIP]{3})$")

    src = os.path.join(_path_dp, razao)
    if not os.path.isdir(src):
        return False, 'Pasta não encontrada'

    result = False, 'Nenhum backup necessario'
    for arq in os.listdir(src):
        match = regex.search(arq)
        if not match:
            continue

        comp = ''.join(i for i in match.group(1) if i.isdigit())
        if comp not in comps:
            continue

        result = True, ''
        dir_src = os.path.join(src, arq)
        dir_dest = os.path.join(_path_bkp, razao)
        os.makedirs(dir_dest, exist_ok=True)
        shutil.copy(dir_src, dir_dest)

        try:
            n_arq = f"SEFIPBKP {comp}.zip"
            os.rename(os.path.join(dir_dest, arq), os.path.join(dir_dest, n_arq))
        except FileExistsError:
            pass

    return result


def get_data_from_csv():
    try:
        file = open(r'ignore\dados.csv', 'r')
        lista = file.readlines()
    except:
        return None, None
    return len(lista), lista


def restaura_bkp(razao, comp):
    print(">>> Restaurando bkp")
    p.press(('alt', 'f', 'r'))
    sleep(1)

    if not click_img('btn_ok.png', conf=0.9, delay=1):
        return '1 - Erro ao restaurar bkp'

    if not wait_img('icon_folder.png', conf=0.9, timeout=30):
        return '2 - Erro ao restaurar bkp'

    p.write(os.path.join(_path_bkp, razao, f"SEFIPBKP {comp}.zip"))
    p.press('enter')

    while True:
        if find_img('bkp_not_found.png', conf=0.9):
            p.press('esc', presses=2)
            return 'Backup não encontrado'
        if find_img('bkp_success.png', conf=0.9):
            p.press('enter')
            return ''
        sleep(1)


def altera_cert():
    print(">>> Alterando certificado")
    cert = ('1', '35586086000160', 'RPEM CONTABIL E GESTAO DE DAD', 'R POSTAL SERVICOS CO')

    p.hotkey('alt', 'l')
    sleep(1)

    for dado in cert:
        sleep(1)
        p.press('del')
        p.write(dado)
        p.press('tab')

    p.press('enter', presses=2, interval=1)
    sleep(2)


def altera_indice():
    print('>>> Alterando indice')
    sleep(1)

    p.press(('enter', 'alt', 'f', 'a', 'enter'), interval=0.5)
    wait_img('icon_folder.png')

    p.write(os.path.join(_path_index, 'Indices.txt'))
    p.press('enter')

    wait_img('icon_close.png')
    p.press('enter', presses=3, interval=0.5)


def abre_mov(comp, venc):
    print('>>> Abrindo movimento')
    while True:
        p.press('f4')
        sleep(1)

        if find_img('cod_rec_115.png', conf=0.9):
            cod_rec = '115'
            break

        if find_img('cod_rec_150.png', conf=0.9):
            cod_rec = '150'
            break

    p.hotkey('alt', 'n')
    while not find_img('open_confirmed.png', conf=0.9):
        if find_img('open_confirm.png', conf=0.9):
            p.hotkey('alt', 's')

    sleep(1)
    p.write(comp)
    p.press('tab')
    p.write(cod_rec)
    sleep(1)
    p.press(('enter', 'tab', 'down', 'tab'))
    p.write(venc.replace('/', ''))
    p.press('enter', presses=2, interval=0.5)

    while not find_img('open_finish.png', conf=0.9):
        if find_img('open_success.png', conf=0.9):
            p.press('enter')
        if find_img('open_erro_disk.png', conf=0.9):
            p.press('enter')
        if find_img('open_erro_indice.png', conf=0.9):
            altera_indice()
    else:
        sleep(0.5)

    p.press('enter')
    sleep(1)
    return ''


def reset_tabs():
    p.hotkey('ctrl', 'm')
    p.press('tab', presses=2)
    p.press(('home', 'down', 'enter', 'tab'), interval=0.5)
    sleep(1)

    if not find_img('icon_func.png', conf=0.9):
        p.hotkey('alt', 'c')
        return True

    p.press('tab', presses=3)
    p.press(('1', 'enter'), interval=0.5)
    p.press('tab', presses=5)
    p.press('enter')
    sleep(1)
    
    p.hotkey('alt', 's')
    if not wait_img('mov_save.png', conf=0.9):
        return False

    p.press('enter')
    return True


def search_func(pis, n_downs):
    print('>>> Procurando funcionario por pis')
    images = {2: 'icon_func_blue.png', 3: 'icon_func_black.png'}
    p.hotkey('ctrl', 'm')
    sleep(1)

    p.press(('tab', 'tab', 'home'))
    p.press('down', presses=n_downs)
    
    sleep(1)
    p.press('enter')
    if not find_img(images[n_downs], conf=0.9):
        p.press('tab')
        p.hotkey('alt', 'c')
        return False

    p.hotkey('shift', 'tab')
    p.write(pis)
    sleep(2)

    for y in range(281, 520, 17):
        if not p.pixelMatchesColor(446, y, (240,) * 3):
            continue
        p.screenshot(r'imgs\tmp.png', region=(446, y, 80, 17))

        with Image.open(r'imgs\tmp.png') as f:
            text = ocr.image_to_string(f)
            text = ''.join(i for i in text if i.isdigit())
            if text == '':
                break

        print(f'\tvalor:{text} controle:{pis}')
        match = [a for a, b in zip(text, pis) if a == b]
        if len(match) < len(text)-2:
            break

        print('\tvalor aceito')
        p.press('tab', presses=6)
        p.press(('home', 'enter'))
        p.press('tab', presses=6)
        p.press('enter')
        p.hotkey('alt', 's')
        if not wait_img('mov_save.png', conf=0.9):
            break

        p.press('enter')
        return True
    else:
        print('')

    p.hotkey('alt', 'c')
    return False


def check_pesquisa(pis):
    if not reset_tabs():
        return False
    return any((search_func(pis, 2), search_func(pis, 3)))


def mov_func(pis):
    print('>>> Buscando funcionario')
    if not click_img('icon_cpart.png', conf=0.9, delay=2, clicks=2):
        return 'Participação não marcada'

    if find_img('icon_func_red.png', conf=0.9):
        if check_pesquisa(pis):
            return ''
        return 'Pesquisa não retornou funcionario'
        
    tmdor = p.locateAllOnScreen(r'imgs\icon_tmdor.png', confidence=0.9)
    if not tmdor:
        return 'Não achou Funcionarios/Tomadores'

    result, tmdor = False, len(list(tmdor)) 
    for i in range(tmdor):
        p.press('down')
        sleep(1)
        if check_pesquisa(pis):
            result = True

    if result:
        return ''
    return 'Pesquisa não retornou funcionario'


def fecha_mov():
    print('>>> Fechando Movimento')
    p.press('home')
    sleep(1)
    p.hotkey('alt', 'c')

    while not find_img('close_btn_finish.png', conf=0.9):
        sleep(1)

        if find_img('icon_close.png', conf=0.9):
            break
        if find_img('erro_incon.png', conf=0.9):
            p.hotkey('alt', 'f')
            return 'inconsistencia'
    else:
        p.write(datetime.today().strftime("%d%m%Y"))
        p.hotkey('alt', 'c')
        if not wait_img('icon_close.png', conf=0.9):
            return 'Erro ao fechar movimento'

    while True:
        p.press('enter')
        if wait_img('close_btn_save.png', conf=0.8) is None:
            continue

        p.hotkey('alt', 's')
        sleep(1)

        if wait_img('close_success.png', conf=0.9):
            break

        # prepara para tentar salvar novamente apos um erro
        if find_img('close_erro_save.png', conf=0.9):
            p.press('enter')
            p.hotkey('alt', 'f')
            p.hotkey('alt', 's')
   
    p.press('enter', presses=3, interval=0.5)
    return ''


def recalc_sefip(razao, pis, comp, venc):
    result = restaura_bkp(razao, comp)
    if result != '':
        return result

    altera_cert()
    result = abre_mov(comp, venc)
    if result != '':
        return result

    print('>>> Marcando participacao')
    if not click_img('icon_spart.png', conf=0.9, delay=1, button='right'):
        return 'Erro ao marcar participacao'
    p.press('down', presses=6)
    p.press('enter')

    if wait_img('icon_close.png', conf=0.9, timeout=10):
        p.press('enter')

    result = mov_func(pis)
    if result != '':
        return result

    result = fecha_mov()
    if result != '':
        return result

    print('>>> Salvando backup')
    p.press(('alt', 'alt', 'f', 'b'))
    if not wait_img('bkp_init.png', conf=0.9):
        return '1 - Erro ao salvar bkp'

    p.press(('left', 'enter'))
    if not wait_img('icon_folder.png', conf=0.9):
        return '2 - Erro ao salvar bkp'

    p.write(os.path.join(_path_dp, razao, f'SEFIPBKP {comp}.zip'))
    p.press('enter')
    if not wait_img('bkp_success.png', conf=0.9, timeout=1000):
        return '3 - Erro ao salvar bkp'

    old = f'SEFIPBKP {comp}.zip'
    new = f'{comp}-{pis}-{datetime.now().strftime("%Y%m%d%H%M%S")}.zip'
    
    old = os.path.join(_path_bkp, razao, old)
    new = os.path.join(_path_bkp, razao, new)
    
    p.press('enter')
    os.rename(old, new)
    return 'Procedimento realizado'


def move_sfp():
    aux = tuple(os.path.join(_path_local, i) for i in os.listdir(_path_local))
    aux = max(aux, key=os.path.getmtime, default='')
    os.makedirs(_path_sfp, exist_ok=True)

    for arq in os.listdir(aux):
        if arq[-4:].lower() != '.sfp':
            continue

        shutil.move(os.path.join(aux, arq), _path_sfp)
        return arq

    return 'Não gerou arquivos .sfp'


@time_execution
def run():
    venc = p.prompt(title='Recalculo sefip', text='Data de Vencimento (dd/mm/aaaa):')
    if not check_venc(venc):
        return True
    if not focus(title='SEFIP -', maximize=False):
        return True

    total, lista = get_data_from_csv()
    if total is None:
        return True
    
    print('>>> Gerando guias com vencimento para', venc, '\n')
    for i, dados in enumerate(lista):
        print('>>> Empresa', i, total)

        nome, pis, comps, razao = dados.split(';')
        comps = comps.replace('/', '').split()
        razao = razao.strip()

        res, msg = check_dir(razao, comps)
        if not res:
            print(f'{razao};{msg}\n')
            escreve_relatorio_csv(f'{razao};{msg}')
            return True

        for comp in comps:
            print('>>>', nome, comp)

            text = recalc_sefip(razao, pis, comp, venc)
            arq = move_sfp()

            print(f'{razao};{nome};{pis};{comp};{text};{arq}\n')
            escreve_relatorio_csv(f'{razao};{nome};{pis};{comp};{text};{arq}')

    return True


if __name__ == '__main__':
    run()
