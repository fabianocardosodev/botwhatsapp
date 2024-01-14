""" 
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES MANDAR MENSGS DE COBRANÇA
"""

#Descrever os passos manuais e dps transformar em código

#criar links personalidos do whatsaap e enviar msgs

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui 

# Ler planilha e guardar informações, nome, telefone e data vencimento

webbrowser.open('https//web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('listacontato.xlsx')
pagina_listacontato = workbook['Plan1']

for linha in pagina_listacontato.iter_rows(min_row=2):
    #nome, telefone , vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome} sabemos que o seu aniversário é dia {vencimento.strftime("%d/%m/%Y")}, parabéns!'
    
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
   
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl , w')
        sleep(5)
    except:
        print('Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
    
    
    
    
    
  