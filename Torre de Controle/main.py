##Elaborado por Antonio de Almeida Junior e Eduardo Silva.

from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import pyautogui

def run_automation():
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get('https://go.maxtrack.com.br/')

    sleep(5)

    try:
        email = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div/div[3]/div/div[1]/div[1]')
        email.click()
        pyautogui.write('eduardo.silva@budel.com.br')
        sleep(3)

        senha = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div/div[3]/div/div[1]/div[2]/div/input')
        senha.click()
        pyautogui.write('@Bdt##6235')
        sleep(3)

        entrar = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div/div[3]/div/div[1]/div[3]/button')
        entrar.click()
        sleep(3)

        selecionar_monitoramento = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div[14]/ul/div[3]/li/a')
        selecionar_monitoramento.click()
        sleep(3)

        selecionar_eventos = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div[14]/ul/div[3]/li/ul/li[2]/a/div/span')
        selecionar_eventos.click()
        sleep(3)

        central_eventos = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[1]/div/div[14]/ul/div[3]/li/ul/li[2]/ul/li/a/div/span')
        central_eventos.click()
        sleep(4)

        # XPath Criticidade
        medio = '/html/body/div/div[2]/div[1]/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div[2]/div/table/tbody/tr[1]/td[7]/div'

        while True:
            try:
                # Tentando encontrar o popup
                popup = driver.find_element(By.XPATH, medio)
                print("Popup encontrado!")
                # ... (Suas ações para o popup) ...
                click_aleatorio = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[2]/div/div[4]/div[2]/div/div[1]/div/div[1]/div/div[1]')
                click_aleatorio.click()
                sleep(2)

                abrir_evento = popup.find_element(By.XPATH, medio)
                abrir_evento.click()
                sleep(2)

                iniciar_tratativa = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[2]/div/div[4]/div[1]/div/div/div[1]/div[1]/div[2]/div[3]/a/i')
                iniciar_tratativa.click()
                sleep(5)

                reiniciar_tratativas = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[1]/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div[1]/div[1]/div[1]')
                reiniciar_tratativas.click()
                sleep(5)
                driver.refresh()
                continue
            except NoSuchElementException:
                # Se o popup não for encontrado, espere e continue procurando
                print("Popup não encontrado. Aguardando...")
                sleep(1)
                driver.refresh()
                sleep(10)
                pyautogui.moveTo(1240,239,duration=0.3)
                pyautogui.moveTo(42,374,duration=0.3)
                continue
    except NoSuchElementException:
        print("Erro: Um ou mais elementos não foram encontrados. Reiniciando a automação...")
        driver.quit()
        run_automation()
        sleep(5)
run_automation()