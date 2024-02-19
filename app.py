import openpyxl
from urllib.parse import quote #link especial = quote
import webbrowser
from time import sleep
import pyautogui
from datetime import datetime

planilha = openpyxl.load_workbook('Dados.xlsx')
pagina_clientes = planilha['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    # Tem que converter a data primeiro e depois converter de novo
    vencimento_final = datetime.strptime(vencimento,"%d/%m/%Y")
 
    mensagem = f"""Olá {nome}, sua conta está sobe suspeita de spam. Por favor, contate imediatamente nossa equipe do WhatsApp.\n
    Número: +55 (11)93223-2687\n
    O prazo para contato é até {vencimento_final.strftime("%d/%m/%Y")}, caso o contato não for feito o seu número será desativado.
    """
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        seta = pyautogui.locateCenterOnScreen('setinha.png')
        sleep(1)    
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(3)
        pyautogui.click(1132,176, duration=2)
        sleep(2)

    except:
        print(f'Não foi possível enviar mensagem para: {nome}')
        with open('erros.txt', 'a', newline = '', encoding = 'utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')