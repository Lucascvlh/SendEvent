from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from datetime import datetime, timedelta
from dotenv import load_dotenv
import pandas as pd
import numpy as np
import pyautogui
import time
import os

load_dotenv()

date_actual = datetime.now() + timedelta(days=7)
date_actual_format = date_actual.strftime("%d/%m/%Y")

pathSheet = 'Eventos.xlsx'
sheet = pd.read_excel(pathSheet, sheet_name='Relatorio', usecols=[0,1,2,3,4,5], engine='openpyxl')
directory = 'Resultados'
if not os.path.exists(directory):
  os.makedirs(directory)

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

driver.maximize_window()

driver.get("https://magalu.brainlaw.com.br/Account/Login?ReturnUrl=%2fHome")
driver.find_element(By.XPATH, '//*[@id="Email"]').send_keys(os.getenv('LOGIN_BL'))
driver.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(os.getenv('PASSWORD_BL') + Keys.ENTER)

url = 'https://magalu.brainlaw.com.br/painel/csc#'
driver.get(url)

#Clicar em enviar envento
wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="btnEventoUncleChan"]')))
driver.find_element(By.XPATH,'//*[@id="btnEventoUncleChan"]').click()

wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="gridEventoUncleChan"]/div/div[4]/div/div/div[3]/div[2]/div/div/div/div[1]/input')))

def close():
  driver.find_element(By.XPATH, '//*[@id="ModalSeller"]/div/div/div[3]/button').click()
  time.sleep(2)
  driver.find_element(By.XPATH, '//*[@id="ModalRascuhoEvento"]/div/div/div[3]/button').click()

def finish(message):
    print(message)
    print('Saindo do sistema...')
    for i in range(3, 0, -1): 
        print(f'{i}...')
        time.sleep(1)

def atualizar_plan(message):
    sheet.at[line,'Status'] = message
    sheet.to_excel(pathSheet, sheet_name='Relatorio', index=False)

for line in range(len(sheet)):
  idSheet = str(sheet.at[line, 'ID'])
  if not np.isnan(sheet.at[line, 'ID']):
    if not pd.isna(sheet.at[line, 'Status']):
      continue

  try:
    driver.find_element(By.XPATH,'//*[@id="gridEventoUncleChan"]/div/div[4]/div/div/div[3]/div[2]/div/div/div/div[1]/input').clear()
    driver.find_element(By.XPATH, '//*[@id="gridEventoUncleChan"]/div/div[4]/div/div/div[3]/div[2]/div/div/div/div[1]/input').send_keys(idSheet)
    time.sleep(2)
    idProcess = driver.find_element(By.XPATH, '//*[@id="gridEventoUncleChan"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[3]')
    idProcess_text = idProcess.text
    if idProcess_text == idSheet:
      event = sheet.at[line, 'EVENTO']
      seller = sheet.at[line, 'SELLER']
      if event == 'Excluir':
        driver.find_element(By.XPATH, '//*[@id="gridEventoUncleChan"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[2]/a').click()
        wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[10]/div/div[3]/button[1]')))
        driver.find_element(By.XPATH, '/html/body/div[10]/div/div[3]/button[1]').click()
        time.sleep(10)
        driver.find_element(By.XPATH, '/html/body/div[10]/div/div[3]/button[1]').click()
        atualizar_plan(f'Excluido id {idSheet} com sucesso')
        print(f'Excluido id {idSheet} com sucesso')
      if event == 'Enviar':
        sellerSend = driver.find_element(By.XPATH, '//*[@id="gridEventoUncleChan"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[6]').text
        if str(seller).lower() == str(sellerSend).lower().strip() or str(sellerSend).lower().strip() == 'não informada':
          driver.find_element(By.XPATH,'//*[@id="gridEventoUncleChan"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]/td[1]/a/i').click()
          time.sleep(10)
          wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="FormularioConfirmacaoRascunhoEvento"]/div/div/div/div[1]/div/div[3]/div/div/div/div/div[1]/div')))
          driver.find_element(By.XPATH,'//*[@id="FormularioConfirmacaoRascunhoEvento"]/div/div/div/div[1]/div/div[3]/div/div/div/div/div[1]/div').click()
          time.sleep(1)
          pyautogui.hotkey('down', 'enter')
          for _ in range(2):
            pyautogui.press('tab')
          pyautogui.hotkey('enter')
          wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="gridSeller"]/div/div[4]/div/div/div[3]/div[3]/div/div')))
          time.sleep(2)
          driver.find_element(By.XPATH, '//*[@id="gridSeller"]/div/div[4]/div/div/div[3]/div[2]/div/div/div/div[1]/input').clear()
          driver.find_element(By.XPATH, '//*[@id="gridSeller"]/div/div[4]/div/div/div[3]/div[2]/div/div/div/div[1]/input').send_keys(seller)
          time.sleep(3)
          tr = 1
          td2 = driver.find_element(By.XPATH, f'//*[@id="gridSeller"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[{tr}]/td[2]').text
          if td2 == "":
            close()
            atualizar_plan(f'Sem seller cadastrado no id {idSheet}')
            print(f'Sem seller cadastrado no id {idSheet}')
            continue
          while td2 != "":
            if str(td2).lower() == str(seller).lower():
              td3 = driver.find_element(By.XPATH, f'//*[@id="gridSeller"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[{tr}]/td[3]').text
              if td3 == "":
                close()
                atualizar_plan(f'Sem referência de seller no id {idSheet}')
                print(f'Sem referência de seller no id {idSheet}')
                break
              driver.find_element(By.XPATH, f'//*[@id="gridSeller"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[{tr}]/td[3]').click()
              time.sleep(5)
              driver.find_element(By.XPATH, '//*[@id="ModalSeller"]/div/div/div[3]/button').click()
              time.sleep(2)
              pyautogui.hotkey('shift','tab')
              pyautogui.typewrite(date_actual_format)
              driver.find_element(By.XPATH, '//*[@id="btnEnviar_ucEnviarEventoUncleChan"]/div/span').click()
              wait.until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="btnEnviar_ucEnviarEventoUncleChan"]/div/span')))
              time.sleep(5)
              driver.find_element(By.XPATH,'/html/body/div[10]/div/div[3]/button[1]').click()
              atualizar_plan(f'Enviado id {idSheet} com sucesso')
              print(f'Enviado id {idSheet} com sucesso')
              break
            tr += 1
            td2 = driver.find_element(By.XPATH, f'//*[@id="gridSeller"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[{tr}]/td[2]').text
        else:
          atualizar_plan('Nome do seller divergente.')
          print('Nome do seller divergente.')
    else:
      atualizar_plan(f'Sem dados para o id {idSheet}')
      print(f'Sem dados para o id {idSheet}')
  except NoSuchElementException:
    continue
finish('Processo finalizado.')
driver.quit()
