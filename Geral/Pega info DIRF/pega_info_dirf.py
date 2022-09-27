import glob, os, time, re

from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir, _escreve_relatorio_csv, _escreve_header_csv


def run():
    documentos = ask_for_dir()
    # Analiza cada pdf que estiver na pasta
    for arq in os.listdir(documentos):
        # print(f'\nArquivo: {arq}')
        # Abrir o pdf
        arq = os.path.join(documentos, arq)

        file1 = open(arq, "r")
        dec = file1.read()
        # print(dec)
        # time.sleep(44)
        ano = re.compile(r'DIRF\|(.+)\|N').search(dec).group(1)
        cnpj = re.compile(r'DECPJ\|(.\d+)\|').search(dec).group(1)
        infos_socios = re.findall(r'BPFDEC\|(.\d+)\|(.+)\|\|(.+)\n.+\nRTPO(.+)\n', dec)
        for socio in infos_socios:
            socio = str(socio).split(',')
            cpf = socio[0].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
            nome = socio[1].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
            valores = socio[3].replace("'", "").replace(",", "").replace("(", "").replace(")", "")

            _escreve_relatorio_csv(f"{cnpj};{nome};{cpf};{valores};{ano}", nome='Valores DIRF')
        
        file1.close()
    
    
if __name__ == '__main__':
    run()
