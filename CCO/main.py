import keyboard
import pyautogui
import openpyxl
from openpyxl.styles import PatternFill
from time import sleep, time

def iniciar_codigo():
    print("Código iniciado!")

    # Aguarda pela tecla F ser pressionada
    keyboard.wait("F")

    sleep(5)
    # Abre o arquivo Excel
    caminho_arquivo_excel = 'O:\\TECNOLOGIA\\25_Automações\\CCO\\dados_cco.xlsx'
    planilha_nome = 'Plan1'  # Nome da planilha que você deseja ler
    workbook = openpyxl.load_workbook(caminho_arquivo_excel)
    planilha = workbook[planilha_nome]

    # Define o tempo máximo de espera para a imagem aparecer (em segundos)
    tempo_max_espera = 10

    # Itera sobre as linhas da planilha
    row_number = 2  # Começa da segunda linha
    while row_number <= planilha.max_row:
        linha = planilha[row_number]

        # Certifique-se de que há elementos suficientes antes de tentar acessar os índices
        if len(linha) >= 7:
            #data, valor, lavagem, placa, cnpj = linha
            TITULO = linha[0].value
            CUSTO = linha[1].value
            AREA = linha[2].value
            TABELA = linha[3].value
            PLACA = linha[4].value
            PRODUTO = linha[5].value
            TOTAL = linha[6].value
            DATA = linha[7].value

            ## código de automação ##

            pyautogui.click(35, 40, duration=0.3)
            pyautogui.click(60, 100, duration=0.3)

            # Escrever
            pyautogui.click(134, 159, duration=0.3)
            pyautogui.write(TITULO)

            # Incluir centro de custo
            pyautogui.click(179, 295, duration=0.3)
            pyautogui.write(str(CUSTO))
            keyboard.press('tab')

            ##RESPONSAVEL
            pyautogui.click(212,344,duration=0.3)
            pyautogui.doubleClick(212,344,duration=0.3)
            keyboard.press('delete')
            pyautogui.write("060079")
            keyboard.press('tab')

            # Área
            pyautogui.click(191, 371, duration=0.3)
            pyautogui.write(str(AREA))
            keyboard.press('tab')

            # Tabela de preço
            pyautogui.click(184,428, duration=0.3)
            pyautogui.write(TABELA)
            keyboard.press('tab')

            # Incluir placa
            pyautogui.click(76, 535, duration=0.3)
            pyautogui.click(469, 276, duration=0.3)
            pyautogui.typewrite(PLACA)
            pyautogui.click(458, 280, duration=0.3)
            keyboard.press('tab')
            pyautogui.click(456, 444, duration=0.3)
            #pyautogui.write(data)
            keyboard.press('tab')
            pyautogui.click(762, 541, duration=0.3)
            pyautogui.click(601, 535, duration=0.3)

            # Aba itens
            pyautogui.click(129, 98, duration=0.3)
            pyautogui.click(65, 287, duration=0.3)
            pyautogui.click(205, 131, duration=0.3)
            pyautogui.click(538, 177, duration=0.3)
            pyautogui.click(474, 231, duration=0.3)

            # Filtro
            pyautogui.click(615, 177, duration=0.3)
            pyautogui.write(PRODUTO)
            pyautogui.click(607, 180, duration=0.3)
            keyboard.press('backspace')
            pyautogui.click(959, 177, duration=0.3)
            pyautogui.click(719, 240, duration=0.3)
            pyautogui.click(832, 614, duration=0.3)

            # Quantidade e total
            pyautogui.click(170, 193, duration=0.3)
            pyautogui.write('1')
            pyautogui.click(567, 194, duration=0.3)
            pyautogui.write(str(TOTAL))
            keyboard.press('tab')

            # Tipo de compra
            pyautogui.click(520, 250, duration=0.3)
            pyautogui.write('s')
            keyboard.press('tab')
            
            # Encerrar processo
            pyautogui.click(318, 281, duration=0.3)
            pyautogui.click(1096,702,duration=0.3)
            sleep(2)

            # Tempo inicial para a verificação da imagem
            tempo_inicial_verificacao = time()

            # Verificar se a imagem está na tela
            while time() - tempo_inicial_verificacao < tempo_max_espera:
                if pyautogui.locateOnScreen('O:\\MANUTENCAO\\LAVAGENS\\Projeto\\confirmar.png', confidence=0.8) is not None:
                    # Se a imagem for encontrada, execute o último click e saia do loop
                    pyautogui.click(883,421, duration=0.3)  
                    sleep(0.3)
                    break
                sleep(1)
            else:
                print("Imagem não encontrada. Encerrando o código.")
                # Se a imagem não for encontrada, encerrar o código
                break

            # Avançar para a próxima linha
            row_number += 1

        else:
            # Avançar para a próxima linha se a linha atual não tiver elementos suficientes
            row_number += 1

    # Salvar as alterações no arquivo Excel
    workbook.save(caminho_arquivo_excel)
    # Fechar o arquivo Excel após terminar de usar
    workbook.close()

# Chamar a função para iniciar o código
iniciar_codigo()