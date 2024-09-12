import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import PySimpleGUI as sg

# Interface
sg.theme('Reddit')
layout = [
    [sg.Text('Escolha a planilha para enviar as mensagens')],
    [sg.Input(), sg.FileBrowse('Escolher Planilha', file_types=(("Excel Files", "*.xlsx"),))],
    [sg.Button('Enviar Mensagens', size=(20,1))],
    [sg.Output(size=(50, 10))]  # área para mostrar os logs
]

# Janela
janela = sg.Window('Bot para Enviar Mensagens', layout)

# Eventos
while True:
    evento, valores = janela.read()

    if evento == sg.WINDOW_CLOSED:
        break

    if evento == 'Enviar Mensagens':
        arquivo_planilha = valores[0]
        if not arquivo_planilha:
            print("Por favor, escolha uma planilha.")
            continue

        try:
            # Carregar a planilha
            workbook = openpyxl.load_workbook(arquivo_planilha)
            pagina_clientes = workbook.active  # seleciona a primeira planilha
            
            for linha in pagina_clientes.iter_rows(min_row=2):
                nome = linha[0].value
                telefone = linha[1].value
                vencimento = linha[2].value

                # Formatar a mensagem
                mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link: https://www.link_do_pagamento.com'
                link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                
                # Abrir o WhatsApp Web e enviar a mensagem
                webbrowser.open(link_mensagem_whats)
                sleep(20)  # Ajuste o tempo conforme necessário

                try:
                    # Enviar a mensagem pressionando Enter
                    pyautogui.press('enter')
                    sleep(3)
                    pyautogui.hotkey('ctrl', 'w')  # Fecha a aba do WhatsApp Web
                    sleep(3)

                    print(f'Mensagem enviada para {nome} ({telefone})')
                except Exception as e:
                    print(f'Não foi possível enviar mensagem para {nome} ({telefone}): {e}')
                    with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                        arquivo.write(f'{nome},{telefone}\n')

        except Exception as e:
            print(f'Erro ao processar a planilha: {e}')

janela.close()
