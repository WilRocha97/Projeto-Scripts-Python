import os, time
import shutil
import fitz
import re
from sys import path

path.append(r'..\..\_comum')
from comum_comum import ask_for_dir

cnpjs = ["07.237.006/0001-26",
         "48.339.006/0001-14",
         "33.842.008/0001-54",
         "09.377.019/0001-26",
         "12.793.042/0001-06",
         "00.225.046/0001-09",
         "00.656.064/0001-45",
         "30.324.077/0001-88",
         "06.375.087/0001-68",
         "24.256.088/0001-75",
         "50.926.096/0001-91",
         "42.233.100/0001-08",
         "34.242.100/0001-46",
         "34.194.125/0001-11",
         "08.968.127/0001-00",
         "00.738.128/0001-57",
         "20.129.129/0001-01",
         "02.285.133/0001-87",
         "30.586.137/0001-30",
         "15.163.155/0001-16",
         "07.012.158/0001-20",
         "24.453.167/0001-76",
         "08.386.169/0001-33",
         "29.281.174/0001-51",
         "71.973.176/0001-78",
         "01.216.181/0001-50",
         "57.706.186/0001-16",
         "11.066.188/0001-97",
         "30.064.195/0001-02",
         "35.656.196/0001-51",
         "12.321.221/0001-40",
         "23.577.227/0001-08",
         "54.690.235/0001-81",
         "26.826.236/0001-75",
         "42.259.237/0001-31",
         "01.072.240/0001-64",
         "38.291.242/0001-54",
         "27.655.244/0001-69",
         "65.583.247/0001-23",
         "19.842.258/0001-63",
         "22.958.266/0001-84",
         "66.799.271/0001-67",
         "04.233.287/0001-50",
         "40.381.292/0001-74",
         "10.230.301/0001-65",
         "37.656.316/0001-46",
         "35.272.317/0001-61",
         "29.829.333/0001-00",
         "09.137.339/0001-09",
         "10.380.348/0001-05",
         "08.585.357/0001-90",
         "15.805.357/0001-14",
         "54.136.361/0001-99",
         "17.839.372/0001-09",
         "27.013.373/0001-53",
         "00.709.391/0001-18",
         "42.679.393/0001-51",
         "50.982.396/0001-98",
         "09.254.402/0001-97",
         "33.964.405/0001-07",
         "24.073.410/0001-20",
         "35.455.410/0001-01",
         "10.226.417/0001-20",
         "44.643.427/0001-83",
         "39.775.430/0001-10",
         "04.190.442/0001-07",
         "26.041.443/0001-14",
         "11.714.444/0001-05",
         "04.423.445/0001-35",
         "07.518.452/0001-09",
         "04.709.454/0001-97",
         "17.906.456/0001-18",
         "00.369.460/0001-91",
         "41.666.463/0001-74",
         "04.759.469/0001-60",
         "07.724.474/0001-25",
         "19.840.481/0001-71",
         "15.632.485/0001-03",
         "30.966.496/0001-13",
         "43.039.497/0001-64",
         "03.711.499/0001-33",
         "30.854.501/0001-04",
         "44.305.507/0001-29",
         "11.929.507/0001-40",
         "45.780.509/0001-32",
         "18.121.516/0001-50",
         "04.759.523/0001-77",
         "46.988.523/0001-99",
         "38.080.534/0001-48",
         "08.468.535/0001-01",
         "42.836.535/0001-47",
         "03.893.550/0001-75",
         "04.237.555/0001-02",
         "55.654.560/0001-51",
         "50.980.572/0001-52",
         "10.944.576/0001-60",
         "34.244.577/0001-60",
         "00.811.577/0001-83",
         "40.333.582/0001-42",
         "50.035.583/0001-64",
         "22.528.591/0001-07",
         "08.740.601/0001-42",
         "34.132.604/0001-03",
         "05.805.608/0001-07",
         "54.134.614/0001-95",
         "43.757.621/0001-27",
         "07.245.623/0001-73",
         "54.490.628/0001-41",
         "22.986.631/0001-64",
         "20.014.633/0001-66",
         "54.738.638/0001-53",
         "60.307.642/0001-60",
         "27.112.650/0001-85",
         "22.986.653/0001-24",
         "39.646.667/0001-00",
         "41.075.678/0001-10",
         "59.447.680/0001-39",
         "31.068.682/0001-06",
         "07.589.700/0001-02",
         "30.500.702/0001-03",
         "05.353.705/0001-06",
         "02.188.709/0001-98",
         "39.798.721/0001-24",
         "37.297.726/0001-48",
         "07.812.734/0001-14",
         "41.109.740/0001-48",
         "26.684.750/0001-13",
         "08.022.752/0001-65",
         "17.446.756/0001-61",
         "32.249.759/0001-07",
         "13.379.766/0001-70",
         "17.533.772/0001-91",
         "29.800.775/0001-23",
         "67.072.777/0001-32",
         "19.174.778/0001-45",
         "34.162.779/0001-63",
         "20.486.797/0001-96",
         "57.059.800/0001-03",
         "26.650.800/0001-41",
         "01.855.804/0001-35",
         "09.582.812/0001-67",
         "31.472.815/0001-05",
         "05.346.817/0001-30",
         "54.569.819/0001-01",
         "24.872.820/0001-31",
         "02.446.824/0001-15",
         "29.829.836/0001-85",
         "22.255.847/0001-50",
         "23.649.851/0001-65",
         "17.733.857/0001-13",
         "03.524.880/0001-93",
         "04.434.885/0001-98",
         "31.526.885/0001-90",
         "24.839.885/0001-85",
         "71.623.888/0001-67",
         "07.072.907/0001-05",
         "35.776.927/0001-00",
         "04.672.934/0001-20",
         "09.423.935/0001-55",
         "03.033.940/0001-75",
         "67.418.947/0001-98",
         "73.195.950/0001-92",
         "11.207.955/0001-30",
         "21.975.969/0001-58",
         "28.045.996/0001-70",
         "34.242.100/0002-27",
         "02.285.133/0002-68",
         "21.975.969/0002-39",
         "02.285.133/0003-49",
         "08.585.357/0004-33",
         ]

# palavra pesquisada"
# palavra = input("Qual palavra deseja pesquisar? ")
# re.compile(r'CP-SEGUR.\n11/2021')

# pasta final
pasta = os.path.join('documentos', 'resultados')
os.makedirs(pasta, exist_ok=True)
documentos = ask_for_dir()

# for cnpj in cnpjs:
for file in os.listdir(documentos):
    if not file.endswith('.pdf'):
        continue
    arq = os.path.join(documentos, file)
    doc = fitz.open(arq, filetype="pdf")
    
    for page in doc:
        texto = page.get_text('text', flags=1 + 2 + 8)
        # print(texto)
        # time.sleep(55)
        # regex_termo = re.compile(r'{}'.format(cnpj))
        # regex_termo = re.compile(r'Não foram detectadas pendências/exigibilidades suspensas nos controles da Receita Federal e da Procuradoria-Geral da Fazenda Nacional.')
        # regex_termo = re.compile(r'CP-SEGUR.\n10/2021')
        # regex_termo = re.compile(r'GIA ST-1/1')
        regex_termo = re.compile(r'GIA-1/1')
        # regex_termo = re.compile(r'GIA')
        resultado = regex_termo.search(texto)
        if not resultado:
            continue
        else:
            doc.close()
            new = os.path.join(pasta, file)
            shutil.move(arq, new)
            break
