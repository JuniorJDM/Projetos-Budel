##Elaborado por Eduardo Silva.

import pyautogui
import time

# Automação ALT + TAB
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')

imagem = "C:\\Users\\eduardo.silva\\Desktop\\Robo\\Robo Fadiga\\Screenshot_4.png"

# Número máximo de tentativas
tentativas_maximas = 721

while True:
    for tentativa in range(1, tentativas_maximas + 1):
        try:
            acharimagem = pyautogui.locateOnScreen(imagem)
        except Exception as e:
            print(f"Erro ao tentar localizar a imagem: {e}")
            acharimagem = None

        if acharimagem is not None:
            print(f"Imagem encontrada na tentativa {tentativa}")
            pyautogui.click(91, 359, duration=0.5)
            pyautogui.mouseDown()
            pyautogui.moveTo(591, 359, duration=0.5)
            pyautogui.hotkey('ctrl', 'c')
            pyautogui.mouseUp

            time.sleep(1)

            pyautogui.keyDown('alt')
            pyautogui.hotkey(['tab', 'tab'])
            pyautogui.keyUp('alt')

            time.sleep(1)

            pyautogui.hotkey('ctrl', 'v')

            time.sleep(1)

            pyautogui.hotkey('Enter')

            time.sleep(1)

            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            pyautogui.keyUp('alt')

            time.sleep(1)

            pyautogui.click(1338, 699, duration=1)

            pyautogui.keyDown('alt')
            pyautogui.hotkey('tab', 'tab')
            pyautogui.keyUp('alt')


            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            pyautogui.keyUp('alt')

            time.sleep(20)

            break  # Sai do loop de tentativas após realizar as ações para uma imagem encontrada

        else:
            print(f"Imagem não encontrada na tentativa {tentativa}")
            pyautogui.click(1338, 699, duration=1)
            time.sleep(20)  # Aguarda antes da próxima tentativa

    # Após sair do loop de tentativas, continua procurando (reinicia o loop while)
    # Não há necessidade de verificar imagem_encontrada aqui porque sempre tentará novamente

    print("Fim do processo")