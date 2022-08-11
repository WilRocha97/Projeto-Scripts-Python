# -*- coding: utf-8 -*-
from sys import path
import os, re, sys

path.append(r'..\..\_comum')
from comum_comum import escreve_relatorio_csv


def analisa_arquivo(arq):
    regex_nome = re.compile(r'((\w+\s)*)\s*')
    regex_pasep = re.compile(r'(\d*)\s\d*')
    regex_emp = re.compile(r'((\w+\s)*)\s+\d*')
    regex_comp = re.compile(r'(\d{2}\/\d{4})\s*')
    
    with open(arq, 'r') as file:
        texto = file.readlines()

    nome, pasep, empresa, comps = ('','') * 2
    dados = {}

    for i, linha in enumerate(texto):
        linha = linha.strip()
        if not linha: continue
        
        if linha.startswith("FGTS - EXTRATO DE CONTA"):
            nome, pasep, empresa, comps = ('','') * 2
        elif linha.startswith("NOME DO TRABALHADOR"):
            nome = regex_nome.search(texto[i+1].strip())
        elif linha.startswith("PIS/PASEP"):
            pasep = regex_pasep.search(texto[i+1].strip())
        elif linha.startswith("NOME DO EMPREGADOR"):
            empresa = regex_emp.search(texto[i+1].strip())
        elif linha.startswith("COMPETENCIAS NAO LOCALIZADAS"):
            comps = []
            for aux in texto[i+1:]:
                comp = regex_comp.findall(aux.strip())
                if not comp: break
                comps += comp

        if not all((nome, comps, pasep, empresa)):
            continue

        empresa = empresa.group(1).strip(" ")
        nome = nome.group(1).strip(" ")
        comps = " ".join(comps)
        pasep = pasep.group(1)
        if dados.get(pasep):
            dados[pasep][2] += " " + comps
        else:
            dados[pasep] = [nome, pasep, comps, empresa]

        nome, pasep, empresa, comps = ('','') * 2
                
    for values in dados.values():
        print(values, '\n')
        escreve_relatorio_csv(';'.join(values))

    return True


if __name__ == '__main__':
    for arq in os.listdir('ignore'):
        if not arq.endswith('.txt'): continue
        analisa_arquivo(os.path.join('ignore', arq))
