import pdfplumber
import re, os
from datetime import datetime, date

_RELATORIO_DPCUCA = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Relatorios', 'pdf_dp')
_RELATORIO_DPCUCA_FINAL = os.path.join(r'\\VPSRV02', 'DCA', 'Setor Robô', 'Relatorios', 'dp')
_REGEX_DPCUCA_PDF = re.compile(r'^dpcuca\s-\s(\d{2})(\d{2})(\d{4})\.[pdfPDF]{3}$')


def select_arq():
    data = date(1970, 1, 1)

    for arq in os.listdir(_RELATORIO_DPCUCA):
        match = _REGEX_DPCUCA_PDF.search(arq)
        if match:
            match = [int(i) for i in match.groups()]
            data_arq = date(match[2], match[1], match[0])
            if data_arq > data:
                to_use_arq, data = arq, data_arq

    return os.path.join(_RELATORIO_DPCUCA, to_use_arq)


def gera_rel_dpcuca():
    arq = select_arq()
    regex = re.compile(r'^\d+\s\|\s\d{2}')

    nome_file = 'Relatorio DPCUCA ' + datetime.now().strftime('%d%m%Y') + '.csv'
    file = open(os.path.join(_RELATORIO_DPCUCA_FINAL, nome_file), 'w')

    with pdfplumber.open(arq) as pdf:
        print("Criando planilha...")
        for pagina in pdf.pages[:]:
            texto = pagina.extract_text()
            if not texto: continue
            lines = [line.replace(' | ', ';') + '\n' for line in texto.split('\n') if regex.search(line)]
            file.writelines(lines)
        print("✔ Planilha criada...")
        file.close()


def run():
    comeco = datetime.now()
    print(comeco)

    gera_rel_dpcuca()

    print("Tempo de execução: ", datetime.now() - comeco)


if __name__ == '__main__':
    run()
