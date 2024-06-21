##Elaborado por Antonio de Almeida Junior.
import keyboard
import pyautogui
import openpyxl
from time import sleep

def iniciar_codigo():
    print("Código iniciado!")

    # Aguarda pela tecla F ser pressionada
    keyboard.wait("F")

    # Abra o arquivo Excel
    caminho_arquivo_excel = 'C:\\Users\\aprendiz.pcm\\OneDrive\\Projetos_V2\\O.s_Web\\FORMATO2.xlsx'
    planilha_nome = 'Plan1'  # Nome da planilha que você deseja ler
    workbook = openpyxl.load_workbook(caminho_arquivo_excel)
    planilha = workbook[planilha_nome]

    # Itera sobre as linhas da planilha
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        # Certifique-se de que há elementos suficientes antes de tentar acessar os índices
        if len(linha) >= 5:
            #data, valor, lavagem, placa, cnpj = linha
            DATA = linha[0]
            PLACA = linha[1]
            CNPJ = linha[2]
            LAVAGEM = linha[3]
            VALOR = linha[4]
            #responsavel = linha[5]  # Se necessário

            ## código de automação ##

            pyautogui.click(80, 175, duration=0.5)    
            pyautogui.click(180, 174, duration=0.5)
            pyautogui.click(180, 174, duration=0.5)
            pyautogui.click(73, 222, duration=0.5)
            pyautogui.click(173, 285, duration=0.5)
            pyautogui.write(str(DATA))  # Garanta que a data seja uma string

            #keyboard.press("tab")

            pyautogui.click(403, 306, duration=0.5)
            pyautogui.click(204, 440, duration=0.5)
            pyautogui.click(177, 326, duration=0.5)
            pyautogui.write(str(PLACA))  # Garanta que a placa seja uma string

            keyboard.press("tab")

            pyautogui.click(176, 391, duration=0.5)
            pyautogui.write(str(CNPJ))  # Garanta que o CNPJ seja uma string
            pyautogui.click(174, 414, duration=0.5)
            pyautogui.write("66007")

            keyboard.press("tab")

            pyautogui.click(375, 220, duration=0.5)
            pyautogui.click(102, 247, duration=0.5)
            pyautogui.write(str(LAVAGEM))  # Garanta que a lavagem seja uma string

            keyboard.press("tab")

            pyautogui.click(299, 264, duration=1.0)
            pyautogui.click(217, 275, duration=0.5)
            pyautogui.click(113, 286, duration=0.5)
            pyautogui.write("1")
            pyautogui.click(240, 287, duration=0.5)

            keyboard.press("delete")

            pyautogui.doubleClick(227, 286, duration=0.5)
            keyboard.press("delete")
            pyautogui.write(str(VALOR))  # Garanta que o valor seja uma string

            keyboard.press("tab")

            pyautogui.click(119, 303, duration=0.5)
            pyautogui.click(95, 315, duration=0.5)
            pyautogui.click(228, 325, duration=0.5)
            pyautogui.click(1246, 696, duration=0.5)

            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")
            pyautogui.click(807, 265, duration=1.0)
            keyboard.press("esc")

        sleep(0.2)

    # Feche o arquivo Excel após terminar de usar
    workbook.close()

# Chama a função para iniciar o código
iniciar_codigo()


