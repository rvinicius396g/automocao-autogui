from tkinter.constants import Y
from pandas.core.algorithms import quantile
import pyautogui
import pyperclip
import time
import pandas as pd


pyautogui.PAUSE = 1


#Passo 0 - Abrir o navegador
pyautogui.hotkey("win")
pyautogui.write("Google")
pyautogui.press("enter")

#Passo 1 - Entrar no sistema (no nosso caso, entar no Link)
#usamos o pyperclip, pq o link tem caracteres especiais
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl","v") #hotkey = pressiona um conjuno de caracteres juntos
pyautogui.press("enter") #press = pressiona a tecla enter
time.sleep(3)
pyautogui.click(x=408,y=330,clicks=2)
pyautogui.click(x=427,y=335)
pyautogui.click(x=1719,y=171)
pyautogui.click(x=1493,y=594)

#Passo 2 - Navegar até o local do relatório
#colocamos r"" para considerar os caracteres especiais
tabela = pd.read_excel(r"C:/Users/rober/Downloads/Vendas - Dez.xlsx") #read_excel ler o exce
#printa coluna = exibe
print(tabela)

faturamento = tabela["Valor Unitário"].sum() #Soma os valores da coluna Valor unitário
quantidade = tabela["Quantidade"].sum()


#Mandando perguntas para e-mail

pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=162,y=164) #click on escrever
pyautogui.click(x=1335,y=529) #click on input form
pyperclip.copy("teste@mail.com")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab") #Pula para assunto
pyautogui.write("teste assunto")
pyautogui.press("tab") #pula para o corpo do e-mail
texto = f"""
Olá,
O faturamento foi de: R${faturamento:.2f}
A quantidade de vendas é: {quantidade}
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
pyautogui.hotkey("ctrl","enter")