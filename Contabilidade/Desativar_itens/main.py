import os
import keyboard
import pyautogui
from time import sleep
from datetime import datetime
import requests

def iniciar_codigo():
    print("Código iniciado!")
    # Aguarda pela tecla F ser pressionada
    keyboard.wait("d")
    # Chama a função para iniciar o código
    processar_dados()

def processar_dados():
    registros = []

    with open('C:\\Users\\aprendiz.pcm\\Desktop\\site\\itens.txt', 'r') as arquivo:
        for index, linha in enumerate(arquivo):
            print("Processando linha", index+1)    
            elementos = [valor.strip() for valor in linha.split(',')]
            if len(elementos) >= 1:
                itens = elementos[0]

                ##automação
                pyautogui.click(589,83,duration=0.3)
                pyautogui.click(486,293,duration=0.3)
                pyautogui.write(itens)
                pyautogui.click(636,291,duration=0.3)
                sleep(2)
                pyautogui.click(854,504,duration=0.3)
                pyautogui.click(460,77,duration=0.3)
                pyautogui.click(568,191,duration=0.3)
                pyautogui.click(538,210,duration=0.3)
                pyautogui.click(766,708,duration=0.3)
                pyautogui.click(677,417,duration=0.3)

iniciar_codigo()