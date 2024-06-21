##Elaborado por Antonio de Almeida Junior.

import keyboard
import pyautogui
from time import sleep

def iniciar_codigo():
    print("Código iniciado!")

    #Aguarda pela tecla F ser pressionada
keyboard.wait("F")

    #Chama a função para iniciar o código
iniciar_codigo()

with open('C:\\Users\\aprendiz.pcm\\OneDrive\\Projetos_V2\\S.c_Corp\\dados2.txt', 'r') as arquivo:
    # Enumerate permite obter o índice da linha
    for index, linha in enumerate(arquivo):
        # Certifique-se de que há elementos suficientes antes de tentar acessar os índices
        elementos = [valor.strip() for valor in linha.split(',')]
        if len(elementos) >= 8:
            titulo = elementos[0]
            custo = elementos[1]
            area = elementos[2]
            lavador = elementos[3]
            placa = elementos[4]
            data = elementos[5]
            produto = elementos[6]
            total = elementos[7]
             
           # iniciar processo
            pyautogui.click(35, 40, duration=0.3)
            pyautogui.click(60, 100, duration=0.3)

            # escrever
            pyautogui.click(134, 159, duration=0.3)
            pyautogui.write(titulo)

            # incluir centro de custo
            pyautogui.click(179, 295, duration=0.3)
            pyautogui.write(custo)
            keyboard.press('tab')


            # area
            pyautogui.click(191, 371, duration=0.3)
            pyautogui.write(area)
            keyboard.press('tab')

            # tabela de preço
            pyautogui.click(373, 400, duration=0.3)
            pyautogui.write(lavador)
            keyboard.press('tab')

            # incluir placa
            pyautogui.click(76, 535, duration=0.3)
            pyautogui.click(469, 276, duration=0.3)
            pyautogui.typewrite(placa)
            pyautogui.click(458, 280, duration=0.3)
            keyboard.press('tab')
            pyautogui.click(456, 444, duration=0.3)
            pyautogui.write(data)
            keyboard.press('tab')
            pyautogui.click(762, 541, duration=0.3)
            pyautogui.click(601, 535, duration=0.3)

            # aba itens
            pyautogui.click(129, 98, duration=0.3)
            pyautogui.click(65, 287, duration=0.3)
            pyautogui.click(205, 131, duration=0.3)
            pyautogui.click(538, 177, duration=0.3)
            pyautogui.click(474, 231, duration=0.3)

            # filtro
            pyautogui.click(615, 177, duration=0.3)
            pyautogui.write(produto)
            pyautogui.click(607, 180, duration=0.3)
            keyboard.press('backspace')
            pyautogui.click(959, 177, duration=0.3)
            pyautogui.click(719, 240, duration=0.3)
            pyautogui.click(832, 614, duration=0.3)

            # quantidade e total
            pyautogui.click(170, 193, duration=0.3)
            pyautogui.write('1')
            pyautogui.click(567, 194, duration=0.3)
            pyautogui.write(total)
            keyboard.press('tab')

            # tipo de compra
            pyautogui.click(520, 250, duration=0.3)
            pyautogui.write('l')
            keyboard.press('tab')
            
            # encerrar processo
            pyautogui.click(318, 281, duration=0.3)
            pyautogui.click(1104, 708, duration=0.3)
            pyautogui.click(900, 422, duration=0.3)
            pyautogui.click(900, 422, duration=0.3)
            sleep(0.3)
    




