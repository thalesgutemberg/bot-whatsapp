""" automizar as msg para meus clientes no whatsapp
primeiro passo: onde esta sendo feito? (whatsapp web)
quais tec preciso para resolver essas demandas? - teclado(Pyautogui), -acesso ao site(webbrowser), -automatizar digitação(link zap),
-automaztizar leitura de dados (biblioteca openpyxcl)"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(15)
# ler planilha e guardar informações sobre nome, telefone 
workbook = openpyxl.load_workbook('lista_ficiticia.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=1):
    #nome, telefone
    nome = linha[0].value
    telefone = linha[1].value
    
    mensagem = f'Olá {nome} Você está recebendo uma mensagem de teste do Bot Whatsapp desenvolvido por Thales Araújo'
    
    # criar links personalizados do zap e enviar msg para cada cliente  com base nos dados da planilha
    link_mensagem_zap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_zap)
    sleep(10)
    try:
        seta = pyautogui.locateCenterOnScreen('seta_envio_zap.png', confidence=0.8)
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')