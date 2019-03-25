# pip install openpyxl
# python3 -m pip install openpyxl
from openpyxl import load_workbook

#lendo o arquiv
f = load_workbook('minha_planilha.xlsx')
#lendo a planilha dentro do arquivo
item = f['Plan1']
#delcaracao de array para receber os valores da planilha
values = []
# detecta o final do arquivo (ultima linha)
# depende do cabelcalho da planilha
# TODO testar
eof = 3
while True:
    x = item['A' + str(eof)].value
    if x is None:
        break
    eof += 1

for i in range(3, eof):
    for j in range(ord('A'), ord('Z')+1):
        # armazena todos os valores de uma linha em um vetor
        if item[chr(j)+str(i)].value == '(+)':
            values.append(True)
        elif item[chr(j)+str(i)].value == '(-)':
            values.append(False)
        else:
            num = item[chr(j)+str(i)].value
            num = num[3:]
            values.append(num)
    # criar a string SQL de insercao com os valores de values[]
    sql = "INSERT INTO NNNNNN (campo) VALUES (\'values[0]\')"
    # conectar e inserir no banco
    collection.insert_one(sdf)

    values[:] = []
