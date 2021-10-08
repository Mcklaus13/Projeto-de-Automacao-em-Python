


''' Projeto de automação de envios de e-email.
Projeto criado para automatizar envios de relatórios por e-mail diário,semanal,mensal ou anual.
A linguagem utilizada nesse projeto foi Python e a ferramenta Jupyter Notebook.As bibliotecas foram:Pandas,Pyautogui,time,Pyperclip.
'''


import pyautogui 
import pyperclip
import time
pyautogui.PAUSE = 1 

# Passo 1: Entrar no sistema 
pyautogui.hotkey("ctrl","t")
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing\n"')
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")

time.sleep(5)
# Passo 2: Navegar até o local do relatório
pyautogui.click(x=329, y=277,clicks=2)
time.sleep(2)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=344, y=270,clicks=1)
time.sleep(2)
pyautogui.click(x=759, y=177,clicks=1)
time.sleep(2)
pyautogui.click(x=560, y=564,clicks=1)
time.sleep(5)


# Passo 4: Calcular os indicadores Faturamento e Qtd de produtos
import pandas as pd
 
tabela = pd.read_excel(r"C:\Users\victo\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()



# Passo 5: Entrar no email

pyautogui.hotkey("ctrl","t")
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
time.sleep(5)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=102, y=223,clicks=1)
time.sleep(3)

pyautogui.write('victormcklaustreze@gmail.com')
time.sleep(1)
pyautogui.press("tab")
time.sleep(2)
pyautogui.press("tab")
time.sleep(2)
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl','v')
time.sleep(1)
pyautogui.press('tab')
texto = f'''
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de:{quantidade:,}

Abs
Victor Mcklaus'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','enter')


# time.sleep(5)
# pyautogui.position() Comando para saber a posição do mouse na tela ( no caso a do meu computador)

