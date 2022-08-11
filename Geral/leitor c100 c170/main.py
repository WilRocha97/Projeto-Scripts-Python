

file =  open('Dados.txt', 'r', encoding='utf-8')
dados = file.readlines()
file.close()

with open('resultado.txt', 'w', encoding='utf-8') as resultado:
    filtered = [linha for linha in dados if linha[1:5] in ['C100', 'C170']]
    for i, dado in enumerate(filtered):
        lista = dado.split('|')
        if lista[1] == 'C100':
            chave = lista[9]
        else:
            lista[12] = chave
            filtered[i] = "|".join(lista)
    resultado.writelines(filtered)
    