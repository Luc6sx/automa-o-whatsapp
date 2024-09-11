import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


workbook = openpyxl.load_workbook('Pasta1.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row = 2):
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome} seu boletovence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link: https://www.link_do_pagamento.com'
    link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    #https://web.whatsapp.com/send?phone=&text=
    webbrowser.open(link_mensagem_whats)
    sleep(15)
    try:
        
        seta = pyautogui.locateCenterOnScreen('image.png')
        sleep(3)
        pyautogui.click(seta[0], seta[1])
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(3)
        
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv','a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome} {telefone} \n')

