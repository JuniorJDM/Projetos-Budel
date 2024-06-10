import keyboard
import pyautogui
import openpyxl
from time import sleep

def iniciar_codigo():
    print("Código iniciado!")

    # Aguarda pela tecla F ser pressionada
    keyboard.wait("F")

    # Abra o arquivo Excel
    caminho_arquivo_excel = 'O:\\TECNOLOGIA\\25_Automações\\dados_cco.xlsx'
    planilha_nome = 'Plan1'  # Nome da planilha que você deseja ler
    workbook = openpyxl.load_workbook(caminho_arquivo_excel)
    planilha = workbook[planilha_nome]

    # Itera sobre as linhas da planilha
    for linha in planilha.iter_rows(min_row=2, values_only=True):
        # Certifique-se de que há elementos suficientes antes de tentar acessar os índices
        if len(linha) >= 7:
            #data, valor, lavagem, placa, cnpj = linha
            TITULO = linha[0]
            CUSTO = linha[1]
            AREA = linha[2]
            TABELA = linha[3]
            PLACA = linha[4]
            PRODUTO = linha[5]
            TOTAL = linha[6]
            DATA = linha[7]

            ## código de automação ##

            pyautogui.click(35, 40, duration=0.3)
            pyautogui.click(60, 100, duration=0.3)

            # escrever
            pyautogui.click(134, 159, duration=0.3)
            pyautogui.write(TITULO)

            # incluir centro de custo
            pyautogui.click(179, 295, duration=0.3)
            pyautogui.write(str(CUSTO))
            keyboard.press('tab')

            ##RESPONSAVEL
            pyautogui.click(212,344,duration=0.3)
            pyautogui.doubleClick(212,344,duration=0.3)
            keyboard.press('delete')
            pyautogui.write("060079")
            keyboard.press('tab')

            # area
            pyautogui.click(191, 371, duration=0.3)
            pyautogui.write(str(AREA))
            keyboard.press('tab')

            # tabela de preço
            pyautogui.click(184,428, duration=0.3)
            pyautogui.write(TABELA)
            keyboard.press('tab')

            # incluir placa
            pyautogui.click(76, 535, duration=0.3)
            pyautogui.click(469, 276, duration=0.3)
            pyautogui.write(PLACA)
            pyautogui.click(458, 280, duration=0.3)
            keyboard.press('tab')
            pyautogui.click(456, 444, duration=0.3)
            #pyautogui.write(data)
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
            pyautogui.write(PRODUTO)
            pyautogui.click(607, 180, duration=0.3)
            keyboard.press('backspace')
            pyautogui.click(959, 177, duration=0.3)
            pyautogui.click(719, 240, duration=0.3)
            pyautogui.click(832, 614, duration=0.3)

            # quantidade e total
            pyautogui.click(170, 193, duration=0.3)
            pyautogui.write('1')
            pyautogui.click(567, 194, duration=0.3)
            pyautogui.write(str(TOTAL))
            keyboard.press('tab')

            # tipo de compra
            pyautogui.click(520, 250, duration=0.3)
            pyautogui.write('s')
            keyboard.press('tab')
            
            # encerrar processo
            pyautogui.click(318, 281, duration=0.3)

            #FORMA DE PAGAMENTO
            pyautogui.click(224,98,duration=0.3)
            pyautogui.click(153,190,duration=0.3)
            pyautogui.write(str(DATA))
            keyboard.press('tab')
            pyautogui.click(260,217,duration=0.3)
            pyautogui.write('b')
            keyboard.press('tab')
            pyautogui.click(359,220,duration=0.3)
            pyautogui.click(1095,700,duration=0.3)
            sleep(2)
            pyautogui.click(894,420,duration=0.3)
            sleep(0.3)

    # Feche o arquivo Excel após terminar de usar
    workbook.close()

# Chama a função para iniciar o código
iniciar_codigo()
