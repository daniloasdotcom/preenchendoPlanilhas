import openpyxl
import os
os.getcwd()
path = "C:\\Users\\Usuário\\Desktop"
os.chdir(path)

# primeiro capituramos as informaçãoes básicas para construirmos nossa coluna de dados
col_input = str(input('Qual coluna você deseja preencher?: ')).upper()
num_de_cell = int(input('quantas celulas você deseja preencher na coluna A?: ')) + 1

# Antecipamos uma lista vazia que receberá as coordenadas das celulas que queremos contruir
lista_de_celulas = []

# O laço for será responsável por definir o número de celulas que receberá os dados
for i in range(1, num_de_cell, 1):
    col = col_input
    col = col + str(i)
    lista_de_celulas.append(col)
    print(col)
    print(lista_de_celulas)

# Criamos o novo arquivo excel
wb = openpyxl.Workbook()
# Nomeamos a planilha que desejamos usar
sheet = wb.get_sheet_by_name('Sheet')
# Por fim salvamos o arquivo pela primeira vez com o nome que escolhermos entre parenteses
wb.save('danilo.xlsx')

# para facilitar nosso código nós renomeamos a lista preenchida no laço anterior
lc = lista_de_celulas

# O laço for abaixo se encarrega de preencher as celulas com uma sequência de números
for n, m in zip(range(1, (len(lc)+1), 1), lc):
    cell = n
    sheet[m] = cell
    wb.save('danilo.xlsx')