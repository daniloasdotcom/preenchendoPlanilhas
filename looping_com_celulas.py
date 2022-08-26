import openpyxl
import os
os.getcwd()
path = "C:\\Users\\Usuário\\Desktop"
os.chdir(path)

# Criamos o novo arquivo
wb = openpyxl.Workbook()
# Nomeamos a planilha que desejamos usar
sheet = wb['Sheet']
# Por fim salvamos o arquivo pela primeira vez
wb.save('danilo.xlsx')

lista_de_dados = [10, 11, 12, 13]
lista_de_celulas = ['A1','A2','A3','A4']

for n, m in zip(lista_de_dados, lista_de_celulas):
    cell = n
    sheet[m] = cell
    wb.save('danilo.xlsx')