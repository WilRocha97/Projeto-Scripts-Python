# -*- coding: utf-8 -*-
import fitz, re, os, pyautogui as p
from datetime import datetime
from dados import arquivos  # os dados contem os nomes dos arquivos que serão separados


def guarda_info(page, matchtexto_nome):
    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, pagina1, pagina2, andamento, arquivo):
    if prevtexto_nome == quem:
        with fitz.open() as new_pdf:
            # Define o nome do arquivo
            nome = prevtexto_nome
            nome = nome.replace('/', ' ')
            text = '{} - {}.pdf'.format(arquivo, nome)

            # Define o caminho para salvar o pdf
            text = os.path.join('Separados DARFs', text)

            # Define a página inicial e a final
            new_pdf.insertPDF(pdf, from_page=pagina1, to_page=pagina2)

            new_pdf.save(text)
            print(nome + andamento)

    prevpagina = page.number
    prevtexto_nome = matchtexto_nome
    return prevpagina, prevtexto_nome


def separa():
    # Para cada arquivo nos dados
    for count, arquivo in enumerate(arquivos, start=1):

        # indice de qual andamento está
        indice = '[ {} de {} ]'.format(str(count), str(len(arquivos)))
        print('\n{} [ {} ]'.format(indice, arquivo))

        # Abrir o pdf
        with fitz.open(r'PDFs\{}.pdf'.format(arquivo)) as pdf:
            # Definir os padrões de regex
            padraozinho_nome = re.compile(r'(.+)\nPIS')
            prevpagina = 0
            paginas = 0

            # Pra cada pagina do pdf
            for page in pdf:
                try:
                    # Pega o texto da pagina
                    textinho = page.getText('text', flags=1 + 2 + 8)
                    # Procura o nome da empresa no texto do pdf
                    matchzinho_nome = padraozinho_nome.search(textinho)
                    if not matchzinho_nome:
                        prevpagina = page.number
                        continue

                    # Guardar o nome da empresa
                    matchtexto_nome = matchzinho_nome.group(1)

                    # Se estiver na primeira página, guarda as informações
                    if page.number == 0:
                        prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                        continue

                    # Se o nome da página atual for igual ao da anterior, soma mais um no indice de paginas
                    if matchtexto_nome == prevtexto_nome:
                        paginas += 1
                        # Guarda as informações da pagina atual
                        prevpagina, prevtexto_nome = guarda_info(page, matchtexto_nome)
                        continue

                    # Se for diferente ele separa a página
                    else:
                        # Se for mais de uma página entra aqui
                        if paginas > 0:
                            # define qual é a primeira página e o nome da empresa
                            paginainicial = prevpagina - paginas
                            andamento = ' - Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1)
                            prevpagina, prevtexto_nome = cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, paginainicial, prevpagina, andamento, arquivo)
                            paginas = 0
                        # Se for uma página entra a qui
                        elif paginas == 0:
                            andamento = ' - Pagina = ' + str(prevpagina + 1)
                            prevpagina, prevtexto_nome = cria_pdf(page, matchtexto_nome,  prevtexto_nome, pdf, prevpagina, prevpagina, andamento, arquivo)
                except:
                    print('❌ ERRO')
                    continue

            # Faz o mesmo dos dois de cima apenas para a(as) ultima(as) página(as)
            if paginas > 0:
                paginainicial = prevpagina - paginas
                andamento = ' - Paginas = ' + str(paginainicial + 1) + ' até ' + str(prevpagina + 1)
                cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, paginainicial, prevpagina, andamento, arquivo)
            elif paginas == 0:
                andamento = ' - Pagina = ' + str(prevpagina + 1)
                cria_pdf(page, matchtexto_nome, prevtexto_nome, pdf, prevpagina, prevpagina, andamento, arquivo)

        # indice de quantos andamentos faltam
        print('[ {} Restantes ]\n'.format((len(arquivos)) - count))


if __name__ == '__main__':
    os.makedirs('Separados DARFs', exist_ok=True)
    inicio = datetime.now()

    quem = p.prompt(text='Qual funcionário?\n(Escreva exatamente como está no Holerite)', title='Script incrível')
    separa()

    print(datetime.now() - inicio)
