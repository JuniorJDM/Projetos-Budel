## BIBLIOTECAS
import os
import keyboard
import pyautogui
from time import sleep
from datetime import datetime
## BIBLIOTECAS


## CODIGO AUTOMAÇÃO
def iniciar_codigo():
    print("Código iniciado!")
    # Aguarda pela tecla F ser pressionada
    keyboard.wait("d")
    # Chama a função para iniciar o código
    processar_dados()

def processar_dados():
    registros = []

    with open('C:\\Users\\aprendiz.pcm\\OneDrive\\Projetos_V2\\Projetos_Budel\\Contabilidade\\dados.txt', 'r') as arquivo:
        for index, linha in enumerate(arquivo):
            print("Processando linha", index+1)    
            elementos = [valor.strip() for valor in linha.split(',')]
            if len(elementos) >= 1:
                produto = elementos[0]

                pyautogui.click(587,80,duration=0.5)
                pyautogui.click(511,294,duration=0.3)
                pyautogui.write(produto)
                pyautogui.click(632,291,duration=0.3)
                pyautogui.click(564,350,duration=0.3)
                pyautogui.click(860,504,duration=0.3)
                pyautogui.click(467,80,duration=0.3)
                pyautogui.click(627,360,duration=0.3)
                pyautogui.click(530,360,duration=0.3)
                keyboard.press('backspace')
                pyautogui.write('C - CUS')
                keyboard.press('enter')
                pyautogui.click(754,708,duration=0.3)
                pyautogui.click(681,417,duration=0.3)
                registros.append(f"Processado {produto}")
                registros.append(f"Alterado o produto para {'C - CUS'}")

    # Criar o diretório "logs" se não existir
    if not os.path.exists('logs'):
        os.makedirs('logs')

    numero_log = 1
    while os.path.exists(f"logs/log{numero_log}.txt"):
        numero_log += 1

    data_atual = datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    # Adiciona a data do processamento como primeiro registro
    registros.insert(0, f"Data do processamento: {data_atual}")

    # Escreve todos os registros no arquivo de log
    with open(f"logs/log{numero_log}.txt", 'w') as log:
        for registro in registros:
            log.write(registro + '\n')

    print(f"Código concluído. Log salvo em 'logs/log{numero_log}.txt'.")

iniciar_codigo()

##CODIGO AUTOMAÇÃO