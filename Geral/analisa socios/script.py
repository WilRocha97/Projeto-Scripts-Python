import os
import re
import fitz


def ler_pdf():
    for file in os.listdir('.'):
        if not file.lower().endswith('.pdf'): continue
        str_empresa = ''
        str_codigo = ''
        str_socio = ''
        regex_empresa = re.compile(r"(\d{5})\s(.+)\n")
        regex_socio = re.compile(r"\d{4}\n([A-Z\.\s\u00C0-\u00DC]+)\n\d{2}\/\d{2}\/\d{4}\n")
        doc = fitz.open(file, filetype="pdf")
        for page in doc:
            texto = page.getText('text')
            empresa = regex_empresa.search(texto)
            if empresa:
                str_empresa = empresa.group(2)
                str_codigo = empresa.group(1)
            socio = regex_socio.findall(texto)
            if socio:
                arq = open('socios cuca.csv', 'a')
                for str_socio in socio:
                    arq.write(f"{str_codigo};{str_empresa};{str_socio}\n")
                arq.close()

        doc.close()



def main():
    ler_pdf()


if __name__ == '__main__':
    main()
