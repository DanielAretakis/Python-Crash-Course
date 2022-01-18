import time

import pandas as pd
import pyautogui
import pyperclip

pyautogui.PAUSE = 0.5
# Passar para inglês

pyautogui.press('win')
pyautogui.write('Opera')
pyautogui.press('enter')
pyautogui.press('enter')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3)
pyautogui.doubleClick(x=384, y=286)
time.sleep(2)
pyautogui.rightClick(x=384, y=286)
pyautogui.click(x=454, y=725)
time.sleep(5)
pyautogui.press('enter')
tabela = pd.read_excel(r'D:\Programação\Intensivo Python\Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
qtdprod = tabela['Quantidade'].sum()
pyautogui.hotkey('ctrl', 't')
pyautogui.write('https://outlook.live.com/mail/0/')
pyautogui.press('enter')
time.sleep(3)
pyautogui.click(x=166, y=177)
time.sleep(2)
pyautogui.write('daniel.aretakis@hotmail.com')
pyautogui.press('tab')
pyperclip.copy('Intensivo Python')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyperclip.copy('''
{}

quantidade total de produtos {:,}
e o faturamento R$ {:,.2f}'''.format(tabela, qtdprod, faturamento))
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
