import pyautogui
import pyperclip
import time
import pandas as pd


# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write ->  escrever um texto
# pyautogui.press ->  apertar uma tecla

pyautogui.PAUSE=0.5

# Passo 1: Acessar o navegador
pyautogui.click(x=957, y=1048)
pyautogui.click(x=1194, y=961)
pyautogui.click(x=587, y=84)
time.sleep(3)

# Passo 2: Entrar no sistema da empresa (no nosso caso vai ser o link do drive)
pyperclip.copy('https://drive.google.com/drive/u/1/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

# demora alguns segundinhos
time.sleep(4)

# Passo 3: Navegar no sistema e encontrar a base de dados (entrar na pasta exportar)
pyautogui.click(x=435, y=356, clicks=2)
time.sleep(6)

# Passo 3: Exportar/Fazer o Dowload da Base de Dados
pyautogui.click(x=480, y=516) # clicar no arquivo
pyautogui.click(x=1665, y=215) # clicar no Mais Ações
pyautogui.click(x=1367, y=676) # clicar no fazer dowload
time.sleep(6)

# Passo 4: Importar a base de dados para o Python usando pandas as pd

tabela = pd.read_excel(r'C:\Users\Hitalo\Downloads\Vendas - Dez(4).xlsx')
print(tabela)

# Passo 5: Calcular os indicadores

# faturamento - soma da coluna Valor Final
faturamento = tabela['Valor Final'].sum()

# quantidade de produtos - soma da coluna Quantidade
quantidade = tabela['Quantidade'].sum()

print(quantidade)
print(faturamento)

# Passo 6: Enviar um email para a diretoria com o relatório

# abrir o email
pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://mail.google.com/mail/u/1/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)

# clicar no escrever
pyautogui.click(x=144, y=242)
time.sleep(6.5)

# escrever o email destinatario
pyautogui.write('h.jacome03@gmail.com')
time.sleep(0.9)
pyautogui.press('tab') # seleciona o destinatario
pyautogui.press('tab') # passar para o campo de assunto
time.sleep(0.9)

# escrever o assunto
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')

# escrever o corpo do email
texto=f'''
Prezados, bom dia

O faturamento de ontem foi de R${faturamento:,.2f}.
A quantidade de produtos foi de {quantidade:,}.

Abs
Hitalo Jacome.'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')

# enviar o email
#pyautogui.click(x=1174, y=991)
pyautogui.hotkey('ctrl','enter')


#### Use esse código para descobrir qual a posição de um item que queira selecionar

#time.sleep(5)
#pyautogui.position()
