import logging
import openpyxl
import pandas as pd
from time import sleep
from pandas import ExcelFile
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
import os


driver = webdriver.Chrome("C:\\Users\\guiga\\Documents\\PyScripts\\ExternalDrivers\\chromedriver.exe")
path = os.getcwd()
print(path)
EXC = pd.ExcelFile(".\\VirtualEnv\\MyScripts\\TR1.xlsx")
EXC = pd.ExcelFile.parse(EXC)

# ******************************** First Loggin ********************************** #
driver.get('https://compras.fieb.org.br/core/default.aspx?U=637719883608560754')
driver.find_element(By.ID,'ctl00_ctl11_tbxLogin').send_keys('guilherme.freire')
driver.find_element(By.ID,'ctl00_ctl11_tbxSenha').send_keys('meridian99')
driver.find_element(By.ID,'ctl00_ctl11_btnAcessar').click()
sleep(2)
# ******************************************************************************* #
Current_Date = (str(datetime.now()))[0:10]
print(Current_Date)

for i in range(0, len(EXC)):
    # *** Get Part Number and Item description on the sheet *** #
    PartNumber = str(EXC['Part Number'][i])
    Description = str(EXC['Especificação'][i])
    AltPartNumber = str(EXC['Alternativa'][i])
    # ********************************************************* #

    driver.get('https://compras.fieb.org.br/core/empresa/produto/produtoManutencao.aspx?q=R0EFSutZFtEZXvMQhyOKWjrBOU6tYrAx_1ANwAB9haqEl54GYHsvxoKMhTtBD56o3mPcWTWAtwlROjlBpflmlg==')
    driver.find_element(By.ID,'img').click()
    print("\nCurrent Window: " + driver.title + "\nCadastrando item: " + PartNumber)
    parent_window = driver.current_window_handle
    sleep(1)
    child_window = driver.window_handles[1]
    driver.switch_to.window(child_window)

    sleep(1)
    driver.find_element(By.ID,'check_191').click()
    driver.switch_to.window(parent_window)
    sleep(1)
    UnitySelec = driver.find_element(By.ID,'td_cPRODUTO_x_nCdUnidadeMedida')
    UnitySelec.click()
    sleep(0.5)
    driver.find_element(By.XPATH,'//*[@id="_cPRODUTO_x_nCdUnidadeMedida_listbox"]/li[43]').click()

    if(AltPartNumber == 'nan'):
        driver.find_element(By.ID,'_cPRODUTO_x_sDsProduto').send_keys(Description + ' - ' + PartNumber)
    else:
        driver.find_element(By.ID,'_cPRODUTO_x_sDsProduto').send_keys(Description + ' - ' + PartNumber + ' ou ' + AltPartNumber)
    driver.find_element(By.ID,'btnSalvarJS').click()
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    ItemCode = driver.find_element(By.ID,'_cPRODUTO_x_sCdProdutoEmpresa').text
    sleep(2)
    driver.find_element(By.ID,'tabAba11').click()
    sleep(1)
    driver.find_element(By.ID,'check_64').click()
    driver.find_element(By.ID,'ctl00_conteudoBotoes_btnConfirmar').click()
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    print('Item Cadastrado')
    with open ('PartNumber_Log - ' + Current_Date + '.txt', 'a') as log_file:
        log_file.write(str(datetime.now())[0:19] + ' > ')
        log_file.write(PartNumber + ' - Codigo Portal: ' + ItemCode + '\n')
    # Approving item #
    sleep(0.5)
    driver.find_element(By.XPATH,'//*[@id="tabAbas"]/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
    driver.find_element(By.ID,'btnEnviarAprovacao').click() #Send to be approved
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    print('Cadastro Enviado para Aprovação')
    ##################
    sleep(2)
print('Processo Finalizado')

