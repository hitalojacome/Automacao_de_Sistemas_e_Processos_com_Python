import pyautogui
import pyperclip
import time


# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar uma tecla

pyautogui.PAUSE=0.5

# Passo 2: Navegar no sistema e encontrar a base de dados (entrar na pasta exportar)
pyautogui.click(x=957, y=1048)
pyautogui.click(x=1194, y=961)
pyautogui.click(x=587, y=84)
time.sleep(5)

# Passo 1: Entrar no sistema da empresa (no nosso caso vai ser o link do drive)
pyperclip.copy('https://drive.google.com/drive/u/1/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

# demora alguns segundinhos
time.sleep(4)

# Passo 2: Navegar no sistema e encontrar a base de dados (entrar na pasta exportar)
pyautogui.click(x=435, y=762, clicks=2)
time.sleep(6)

# Passo 3: Exportar/Fazer o Dowload da Base de Dados
pyautogui.click(x=454, y=779) # clicar no arquivo
pyautogui.click(x=1668, y=218) # clicar nos 3 pontinhos
pyautogui.click(x=1443, y=678) # clicar no fazer dowload
time.sleep(6)

# Passo 4: Importar a base de dados para o Python
import pandas as pd

tabela = pd.read_excel(r'C:\Users\Hitalo\Downloads\Vendas - Dez.xlsx')
print(tabela)
