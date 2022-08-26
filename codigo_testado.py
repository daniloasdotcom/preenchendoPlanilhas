# Importamos as bibliotecas necessárias
import openpyxl
import os
import speech_recognition as sr
from gtts import gTTS
from playsound import playsound
import send2trash

# Todas as funções abaixo capituram o audio e utilizam a informação para desenvolver a planilha
# ouvir_arquivo () captura o nome que será dado ao arquivo xlsx
def ouvir_arquivo():
    # Habilita o microfone para ouvir o usuario
    microfone = sr.Recognizer()
    with sr.Microphone() as source:
        # Chama a funcao de reducao de ruido disponivel na speech_recognition
        microfone.adjust_for_ambient_noise(source)
        # Avisa ao usuario que esta pronto para ouvir
        print("Pode falar o nome a ser dado ao arquivo: ")
        # Armazena a informacao de audio na variavel
        audio = microfone.listen(source)
    try:
        # Passa o audio para o reconhecedor de padroes do speech_recognition
        arquivo = microfone.recognize_google(audio, language='pt-BR')
        # Após alguns segundos, retorna a frase falada
        print("Eu entendi que você falou: " + arquivo)
    # Caso nao tenha reconhecido o padrao de fala, exibe esta mensagem
    except sr.UnknownValueError:
        print("Não entendi")

    return arquivo

# ouvir_coluna () captura o nome que será dado à coluna A
def ouvir_coluna():
    # Habilita o microfone para ouvir o usuario
    microfone = sr.Recognizer()
    with sr.Microphone() as source:
        # Chama a funcao de reducao de ruido disponivel na speech_recognition
        microfone.adjust_for_ambient_noise(source)
        # Avisa ao usuario que esta pronto para ouvir
        print("Pode falar a letra correspondente à coluna: ")
        # Armazena a informacao de audio na variavel
        audio = microfone.listen(source)
    try:
        # Passa o audio para o reconhecedor de padroes do speech_recognition
        coluna = microfone.recognize_google(audio, language='pt-BR')
        # Após alguns segundos, retorna a frase falada
        print("Eu entendi que você falou: " + coluna)
    # Caso nao tenha reconhecido o padrao de fala, exibe esta mensagem
    except sr.UnknownValueError:
        print("Não entendi")

    return coluna

# ouvir_coluna () captura o nome que será dado à coluna A
def ouvir_nomecoluna():
    # Habilita o microfone para ouvir o usuario
    microfone = sr.Recognizer()
    with sr.Microphone() as source:
        # Chama a funcao de reducao de ruido disponivel na speech_recognition
        microfone.adjust_for_ambient_noise(source)
        # Avisa ao usuario que esta pronto para ouvir
        print("Pode falar o nome da lista de dados: ")
        # Armazena a informacao de audio na variavel
        audio = microfone.listen(source)
    try:
        # Passa o audio para o reconhecedor de padroes do speech_recognition
        nomecoluna = microfone.recognize_google(audio, language='pt-BR')
        # Após alguns segundos, retorna a frase falada
        print("Eu entendi que você falou: " + nomecoluna)
    # Caso nao tenha reconhecido o padrao de fala, exibe esta mensagem
    except sr.UnknownValueError:
        print("Não entendi")

    return nomecoluna

# ouvir_coluna () captura o nome que será dado à coluna A
def ouvir_numcell():
    # Habilita o microfone para ouvir o usuario
    microfone = sr.Recognizer()
    with sr.Microphone() as source:
        # Chama a funcao de reducao de ruido disponivel na speech_recognition
        microfone.adjust_for_ambient_noise(source)
        # Avisa ao usuario que esta pronto para ouvir
        print("Pode falar o número total de dados: ")
        # Armazena a informacao de audio na variavel
        audio = microfone.listen(source)
    try:
        # Passa o audio para o reconhecedor de padroes do speech_recognition
        num_cell = microfone.recognize_google(audio, language='pt-BR')
        # Após alguns segundos, retorna a frase falada
        print("Eu entendi que você falou: " + num_cell)
    # Caso nao tenha reconhecido o padrao de fala, exibe esta mensagem
    except sr.UnknownValueError:
        print("Não entendi")

    return num_cell

# ouvir_x() captura os valores que serão dados à celulas
def ouvir_x():
    # Habilita o microfone para ouvir o usuario
    microfone = sr.Recognizer()
    with sr.Microphone() as source:
        # Chama a funcao de reducao de ruido disponivel na speech_recognition
        microfone.adjust_for_ambient_noise(source)
        # Avisa ao usuario que esta pronto para ouvir
        print("Pode falar o valor: ")
        # Armazena a informacao de audio na variavel
        audio = microfone.listen(source)
    try:
        # Passa o audio para o reconhecedor de padroes do speech_recognition
        x = microfone.recognize_google(audio, language='pt-BR')
        # Após alguns segundos, retorna a frase falada
        print("Eu entendi que você falou: " + x)
    # Caso nao tenha reconhecido o padrao de fala, exibe esta mensagem
    except sr.UnknownValueError:
        print("Não entendi")

    return x

# Após definidas as funções, elas serão chamadas para exercer seu papel
# Iniciamos indicando manualmente, o path onde criaremos nosso arquivo
#################################################################################
os.getcwd()
path = "C:\\Users\\Usuário\\Desktop"
os.chdir(path)
#################################################################################


# As linhas a seguir fazem o programa ficar mais dinâmico, onde a assitente se comunica com
# o usuário indicado qual informação deve ser fornecidada via microfone

#################################################################################
# Primeiro ele pede um nome para ser dado ao arquivo a ser editado
nome_arquivo = str("Danilo qual nome você gostaria de dar ao arquivo?")
tts = gTTS(text = nome_arquivo, lang='pt-br')
tts.save(savefile='nome_arquivo.mp3') #Salva o arquivo de audio
playsound('nome_arquivo.mp3') #Da play ao audio
send2trash.send2trash('nome_arquivo.mp3') #Após ouvido, o aúdio é deletado

# Armazenamos então o nome do arquivo na variável "arquivo"
arquivo = str(ouvir_arquivo() + '.xlsx')

wb = openpyxl.Workbook() # Criamos o novo arquivo
sheet = wb.get_sheet_by_name('Sheet') # Nomeamos a planilha que desejamos usar
wb.save(arquivo) # Por fim salvamos o arquivo pela primeira vez
#################################################################################

#################################################################################
# Utilizamos um input manual para indicarmos a coluna que será editada na planilha
col_input = str(input('Qual a letra correspondente à coluna que você deseja preencher?: '))
col_input = str(col_input).upper()

# "Col" armazena a cordenada da primeira celula da coluna indicada
col = str(col_input + '1')
print('A celula que levará o nome escolhido é a', col)
#################################################################################

#################################################################################
# Agora a assistente pede um nome para ser dado ao cabeçalho da coluna de dados
nome_coluna = str("Diga o nome que deve ser dado à coluna de dados")
tts = gTTS(nome_coluna, lang='pt-br')
tts.save('nome_coluna.mp3') #Salva o arquivo de audio
playsound('nome_coluna.mp3') #Da play ao audio
send2trash.send2trash('nome_coluna.mp3') # Após ouvido, o aúdio da assistente é deletado

nomecoluna = str(ouvir_nomecoluna()) # Tranforma o audio falado pelo usuário em uma string

sheet[col] = nomecoluna
wb.save(arquivo)
#################################################################################

# primeiro capituramos as informaçãoes básicas para construirmos nossa coluna de dados

#################################################################################
#num_de_cell = str('Quantas celulas você deseja preencher na coluna?: ')
#tts = gTTS(num_de_cell, lang='pt-br')
#tts.save('num_de_cell.mp3') #Salva o arquivo de audio
#playsound('num_de_cell.mp3')#Da play ao audio
#send2trash.send2trash('num_de_cell.mp3') #Após ouvido, o aúdio é deletado

num_cell = str(input('Qual a letra correspondente à coluna que você deseja preencher?: '))
print('A total de celulas será de ', num_cell)
#################################################################################

# A seguir os laços for criam a coordenadas das celulas que serão preenchidas
# Armazenam essas coordenadas na lista vazia "lista_de_celulas", previamente criada
#################################################################################
#
lista_de_celulas = []
for i in range(2, (num_cell + 2), 1):
    col = col_input
    col = col + str(i)
    lista_de_celulas.append(col)

# As duas linha abaixo servem apenas para verificar as saidas finais do laço for
lc = lista_de_celulas
print('e as celulas serão', lc)

# O ultimo laço for pede ao usuário que diga o valor a ser digitada para cada celula a ser preenchida
for i in lc:
    nome_dadox1 = str("Diga o valor a ser digitado na celula " + i)
    nome_dadox1 = str(nome_dadox1) # garante a leitura da string pelo gTTS()
    tts = gTTS(nome_dadox1, lang='pt-br')
    tts.save('nome_dadox1.mp3') #Salva o arquivo de audio
    playsound('nome_dadox1.mp3') #Da play ao audio
    send2trash.send2trash('nome_dadox1.mp3') #Após ouvido, o aúdio é deletado

    x = float(ouvir_x()) # Captura o dado capitado pelo microfone e grava como tipo float
    print(x) # Printa para verificar o dado ouvido pelo programa
    sheet[i] = x # O programa lança o dado na coordenada da celula em questão
    wb.save(arquivo) # O arquivo criado é salvo a cada looping for
#################################################################################