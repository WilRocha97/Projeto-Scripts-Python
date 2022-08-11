# -*- coding: utf-8 -*-
from datetime import datetime
from sys import path
import time, pyautogui as p

path.append(r'..\..\_comum')
from cuca_comum import _horario, _login, _fechar, _verificar_empresa, _iniciar, _loc_img
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start


def abrir_os_dois(usuario):
    tipos = ["CUCA", "DPCUCA", ]
    for tipo in tipos:
        _iniciar(tipo, usuario)
    time.sleep(1)


def login(empresa, qual_cuca, log, competencia, ano, execucao):
    cod, cnpj, nome = empresa
    texto = '{};{};{};Empresa não encontrada no {}'.format(cod, cnpj, nome, qual_cuca)
    # Verificação de loging e empresa
    if not _login(empresa, log, qual_cuca, 'Exportação DARF-DCTF', competencia, ano):

        if qual_cuca == 'dpcuca':
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Empresa Inativa']), nome=execucao)
            print('❌ Empresa Inativa')
            time.sleep(0.5)

        return False

    if not _verificar_empresa(cnpj, execucao, texto, qual_cuca):
        if qual_cuca == 'dpcuca':
            p.hotkey('win', 'm')
            time.sleep(0.5)

        return False
    return True


def transferir_info(index, empresas, competencia, ano, execucao):
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _hora_limite = datetime.now().replace(hour=17, minute=25, second=0, microsecond=0)
        # Verifica horário
        if _horario(_hora_limite, 'DPCUCA'):
            abrir_os_dois(usuario)
            p.getWindowsWithTitle('DPCUCA')[0].maximize()

        _indice(count, total_empresas, empresa)

        cod, cnpj, nome = empresa

        p.hotkey('win', 'm')
        time.sleep(0.5)

        p.getWindowsWithTitle('DPCUCA')[0].maximize()

        if not login(empresa, 'dpcuca', 'Codigo', competencia, ano, execucao):
            continue

        # Transferir dados do DPCUCA para o CUCA
        while not _loc_img('CUCAInicial.png'):
            print('aqui')
            p.getWindowsWithTitle('DPCUCA')[0].maximize()
            p.getWindowsWithTitle('DPCUCA')[0].activate()

        while not _find_img('Transferir.png', conf=0.9):
            p.getWindowsWithTitle('DPCUCA')[0].activate()
            p.hotkey('alt', 'c')
            time.sleep(1)
            p.press('t')

        _wait_img('Transferir.png', conf=0.9, timeout=-1)
        time.sleep(2)

        while _find_img('CC.png', conf=0.9):
            if _find_img('SelecioneProgramas.png', conf=0.9):
                p.press('enter')
            _click_img('C-C.png', conf=0.9)
            p.moveTo(1, 1)

        while _find_img('Fed.png', conf=0.9):
            if _find_img('SelecioneProgramas.png', conf=0.9):
                p.press('enter')
            _click_img('Federal.png', conf=0.9)
            p.moveTo(1, 1)

        p.hotkey('alt', 't')
        while not _find_img('EscolherEmpresas.png', conf=0.9):
            time.sleep(1)
            if _find_img('SelecioneProgramas.png', conf=0.9):
                p.press('enter')
                time.sleep(0.5)

                while _find_img('CC.png', conf=0.9):
                    if _find_img('SelecioneProgramas.png', conf=0.9):
                        p.press('enter')
                    _click_img('C-C.png', conf=0.9)
                    p.moveTo(1, 1)

                while _find_img('Fed.png', conf=0.9):
                    if _find_img('SelecioneProgramas.png', conf=0.9):
                        p.press('enter')
                    _click_img('Federal.png', conf=0.9)
                    p.moveTo(1, 1)

                p.hotkey('alt', 't')
            time.sleep(1)
        p.hotkey('alt', 'a')

        while not _find_img('TransferidoComSucesso1.png', conf=0.9):
            time.sleep(1)
            if _find_img('TransferidoComSucesso.png', conf=0.9):
                p.press('enter')

        p.press('enter')
        time.sleep(1)
        p.hotkey('alt', 'f')
        time.sleep(0.5)
        p.hotkey('win', 'm')
        time.sleep(0.5)

        _iniciar('CUCA', usuario)

        # Logar na mesma empresa e competencia para confirmar a transferência se houver
        if not login(empresa, 'cuca', 'CNPJ', competencia, ano, execucao):
            continue
        time.sleep(5)

        # Verificar se existe informações para serem transferidas
        if _find_img('ExisteTransferencia.png', conf=0.9):
            p.hotkey('alt', 'a')
            time.sleep(1)

            x = 0
            while not _find_img('DesejaIncluir.png', conf=0.9):
                x += 1
                if x > 10:
                    p.hotkey('alt', 's')
                    time.sleep(1)
                    break

            x = 0
            while not _find_img('TransferenciaFeita.png', conf=0.9):
                x += 1
                if x > 10:
                    p.hotkey('alt', 's')
                    time.sleep(1)
                    break

            p.hotkey('alt', 'o')
            time.sleep(1)
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Transferência concluída']), nome=execucao)
            print('✔ Transferência concluída')

        elif not _find_img('ExisteTransferencia.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Não transferiu']), nome=execucao)
            print('❌ Não transferiu')

    p.hotkey('win', 'm')


@_time_execution
def run():
    execucao = 'Exportação DARF-DCTF 1'
    competencia = 0

    while competencia < 1 or competencia > 12:
        competencia = p.prompt(text='Qual mês?', title='Script incrível')
        competencia = int(competencia)
    ano = p.prompt(text='Qual ano?', title='Script incrível')
    competencia = str(competencia)
    empresas = _open_lista_dados()

    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None: return False

    p.hotkey('win', 'm')

    abrir_os_dois(usuario)
    transferir_info(index, empresas, competencia, ano, execucao)
    _fechar()


if __name__ == '__main__':
    run()
