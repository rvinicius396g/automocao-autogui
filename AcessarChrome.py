import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

#Passo 1 - Entrar no sistema (no nosso caso, entar no Link)
pyautogui.hotkey("win")
pyautogui.write("chrome")
pyautogui.press("enter")