##Elaborado por Eduardo Silva.

import pyautogui
import time

for i in range(50400):
    time.sleep(1)

    # Automação ALT + TAB
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')

    # Pausa para garantir que a janela tenha foco
    time.sleep(0.5)

    # Clique em (402, 419)
    pyautogui.click(402, 419, duration=1)
    time.sleep(5)
    
    # Clique em (1228, 171)
    pyautogui.click(1228, 171, duration=1)
    time.sleep(5)
    
    # Clique em (1233, 212)
    pyautogui.click(1233, 212, duration=1)
    time.sleep(5)

    # Clique em (416, 322)
    pyautogui.click(416, 322, duration=1)
    time.sleep(5)

    # Clique em (1270, 171)
    pyautogui.click(1270, 171, duration=1)
    time.sleep(5)

    # Clique em (143, 140)
    pyautogui.click(143, 140, duration=1)
    time.sleep(5)

    # Pressiona F5 para atualizar a página
    pyautogui.press('f5')
    time.sleep(60)

    # Volta ao ALT + TAB para reiniciar no loop
    pyautogui.keyDown('alt')
    pyautogui.press('tab')
    pyautogui.keyUp('alt')