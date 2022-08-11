# -*- coding: utf-8 -*-
import os, re
import pdfplumber


def analisa_pdf(arq):
    texto = ''
    with pdfplumber.open(arq) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text() + '\n'

    return texto


def registra_relatorio(texto=''):
    if not texto:
        arquivo = open('resumo.csv', 'w')
        arquivo.write('Cod. Cuca;CNPJ;GFIP (Cuca);Diferenca GFIP;INSS (Cuca);Diferenca INSS;Observacao\n')
    else:
        arquivo = open('resumo.csv', 'a')
        arquivo.write(texto+'\n')
    arquivo.close()


def registra_diferenca(cuca, cnpj, gfip, total, gfip_grf, total_rubrica, obs):
    dif_gfip = float(gfip.replace('.', '').replace(',', '.')) - float(gfip_grf.replace('.', '').replace(',', '.'))
    dif_total = float(total.replace('.', '').replace(',', '.')) - float(total_rubrica.replace('.', '').replace(',', '.'))

    texto = ';'.join([cuca, cnpj, gfip, str(dif_gfip).replace('.', ','), total, str(dif_total).replace('.', ','), obs])
    print(texto)
    registra_relatorio(texto)


def analisa_resumo():
    arquivo_resumo = 'Geral.pdf'
    lista = {}
    titulo = 'Resumo por Eventos do Mês'
    regex_cnpj = re.compile(r'(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}|\d{12})')
    regex_cuca = re.compile(r'(\d{5})([\d\w.\s])+\d{5}-\d{3}')
    regex_gfip = re.compile(r'Empresa:\s*(([0-9\.*,]*\s*)+)')
    regex_total = re.compile(r'Valor a Recolher:\s*([0-9\.?]*,\d{2})')
    with pdfplumber.open(arquivo_resumo) as doc:
        for index, pagina in enumerate(doc.pages, 1):
            texto = pagina.extract_text()
            if titulo not in texto: continue

            cnpj = regex_cnpj.search(texto)
            cuca = regex_cuca.search(texto)
            gfip = regex_gfip.search(texto)
            total = regex_total.search(texto)

            total = total.group(1) if total else '0,00'
            cnpj = cnpj.group(1) if cnpj else ''
            cuca = cuca.group(1) if cuca else ''
            print(cuca, cnpj)
            try:
                gfip = gfip.group(1)
                gfip = [g.strip('\n') for g in gfip.split(' ') if g != '']
                tamanho = len(gfip)
                gfip = gfip[-3] if tamanho >= 3 else gfip[tamanho-1]
            except:
                gfip = '0,00'

            if cnpj and cuca:
                lista[cnpj] = [cuca, total, gfip]
            else:
                print('>>> Falha durante analise da página', index)

    return lista

def analisa_grf(grf):
    regex_grf = re.compile(r'TOTAL\s*A\s*RECOLHER\s*([0-9\.?]*,\d{2})')

    texto_grf = analisa_pdf(grf)
    gfip_grf = regex_grf.search(texto_grf)
    gfip_grf = gfip_grf.group(1) if gfip_grf else '0,00'
    
    return gfip_grf


def analisa_rubrica(rubrica):
    regex_rubrica = re.compile(r'TOTAL\s*A\s*RECOLHER\s*([0-9\.?]*,\d{2}\s*){4}([0-9\.?]*,\d{2})')

    texto_rubrica = analisa_pdf(rubrica)
    total_rubrica = regex_rubrica.search(texto_rubrica)
    total_rubrica = total_rubrica.group(2) if total_rubrica else '0,00'
    
    return total_rubrica


def get_codigo(arq):
    regex_cod = re.compile(r'(\d*)\s?[A-Z]*')
    texto = analisa_pdf(arq)
    cnpj = texto.split('\n')[2]
    cnpj = ''.join([i for i in cnpj if i.isdigit()])
    codigo = regex_cod.search(texto.split('\n')[1]).group(1)

    return codigo, cnpj



# ANALISA OS RELATÓRIOS DE RESUMO GERAL MENSAL (arquivo em lote com todas as empresas); 
# analiticoGRF E Rubrica DE CADA EMPRESA, CRUZANDO OS DADOS E APRESENTANDO A DIFERENÇA 
# ENTRE OS VALORES DE INSS E GFIP ENVIADOS
if __name__ == '__main__':
    registra_relatorio()
    lista_grf = {} 
    lista_rubrica = {}
    regex_grf = re.compile(r'^AnaliticoGRF_(\d{10,14})_\d{8}_\d*\.[pdfPDF]{3}$')
    regex_rubrica = re.compile(r'^Rubrica_(\d{10,14})_\d{8}_\d*\.[pdfPDF]{3}$')

    lista_resumo = analisa_resumo()
    for arq in os.listdir('.'):
        extensao = arq[-4:]
        if extensao not in ['.pdf', '.PDF']: 
            continue

        grf = regex_grf.search(arq)
        if grf:
            n = 11 if len(grf.group(1)) <= 11 else 14
            cnpj = grf.group(1).rjust(n, "0")
            lista_grf[cnpj] = arq

        rubrica= regex_rubrica.search(arq)
        if rubrica:
            n = 11 if len(rubrica.group(1)) <= 11 else 14
            cnpj = rubrica.group(1).rjust(n, "0")
            lista_rubrica[cnpj] = arq

    for cnpj, dados in lista_resumo.items():
        obs = ''
        cuca, total, gfip = dados
        n = 11 if len(cnpj) <= 11 else 14
        cnpj = cnpj.rjust(n, '0')
        arq_grf = lista_grf.get(cnpj, '')
        arq_rubrica = lista_rubrica.get(cnpj, '')
        if arq_grf or arq_rubrica:
            if arq_grf:
                gfip_grf = analisa_grf(arq_grf)
            else:
                gfip_grf = '0,00'
                obs = 'AnaliticoGRF nao encontrado'

            if arq_rubrica:
                total_rubrica = analisa_rubrica(arq_rubrica)
            else:
                total_rubrica = '0,00'
                obs = 'Rubrica nao encontrado'

            registra_diferenca(cuca, cnpj, gfip, total, gfip_grf, total_rubrica, obs)
            print()
        else:
            print(f'{cnpj} - Rubrica e AnaliticoGRF nao encontrado(s)')
            registra_relatorio(f'{cuca};{cnpj};;;;;Rubrica e AnaliticoGRF nao encontrados')

    print('>>> Programa Finalizado.')
    input('Digite qualquer tecla para sair.... ')
    