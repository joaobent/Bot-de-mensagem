import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    if not nome or not telefone or not vencimento:
        print(f"Contato inválido na planilha: {nome}, {telefone}")
        continue  

    mensagem = f'Olá, {nome}! Sou o Assistente virtual de João Pedro Bento. Vim avisar que seu pagamento vence no dia {vencimento.strftime("%d/%m/%Y")}. Fique atento para não ultrapassar a data de vencimento!'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(7)
        sleep(3)
        pyautogui.hotkey('enter')  
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(3)

    except Exception as e:
        print(f'Não foi possível enviar a mensagem para {nome}. ERRO: {e}')
        with open('erros.txt', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone} \n')
        continue
    else:
        print(f'Mensagem enviada para {nome}')