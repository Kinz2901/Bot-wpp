"""
MOTIVO DA CRIAÇÃO
"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento

workbook = openpyxl.load_workbook('Relação.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
  # nome, telefone
  nome = linha[0].value
  telefone = linha[1].value

# Criar links personalizados do whatsapp e enviar mensagens para cada cliente
# com base nos dados da planilha
  mensagem = f'Oi {nome} é o Guilherme, to só testando o bot aqui'
  try:
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    seta = pyautogui.locateCenterOnScreen('enviar-wpp.png')
    sleep(5)
    pyautogui.click(seta[0],seta[1])
    sleep(5)
    pyautogui.hotkey('ctrl', 'w')
    sleep(5)
  except:
    print(f'Não foi possível enviar mensagem para {nome}')
    with open('erros.csv', 'a' ,newline='',encoding='utf-8') as arquivo:
      arquivo.write(f'{nome},{telefone}')

