import logging
from openpyxl import Workbook, load_workbook
import pandas as pd
from time import sleep
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import SessionNotCreatedException
import getpass
from xlrd import open_workbook
from xlutils.copy import copy



VERSION = '1.2.0'
LAST_UPDATE = '07/10/22'



print('______________________________________________________')
print("Automatização de Cadastro do Portal de Compras FIEB")
print('Versão ' + VERSION)
print("Desenvolvido por Guilherme Freire")
print('______________________________________________________\n\n')

user = input("Usuário: ")
passw = getpass.getpass('Senha: ')

try:
    driver = webdriver.Chrome(".\\chromedriver.exe")
except SessionNotCreatedException:
    print('\n\nVersão do driver incompatível!')
    print('Necessário realizar o download do driver com versão adequada ao seu navegador')
    input('[Pressione ENTER para encerrar]')
    exit()

XL_PATH = ".\\Cadastro.xlsx"
EXC = pd.ExcelFile(XL_PATH)
EXC = pd.ExcelFile.parse(EXC)


# ******************************** First Loggin ********************************** #
driver.get('https://compras.fieb.org.br/core/default.aspx?U=637719883608560754')
driver.find_element(By.ID,'ctl00_ctl11_tbxLogin').send_keys(user)
driver.find_element(By.ID,'ctl00_ctl11_tbxSenha').send_keys(passw)
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
    #print("\nCurrent Window: " + driver.title + "\nCadastrando item: " + PartNumber)
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

    if(PartNumber == 'nan'):
        driver.find_element(By.ID,'_cPRODUTO_x_sDsProduto').send_keys(Description)
    else:
        if(AltPartNumber == 'nan'):
            driver.find_element(By.ID,'_cPRODUTO_x_sDsProduto').send_keys(Description + ' - ' + PartNumber)
        else:
            driver.find_element(By.ID,'_cPRODUTO_x_sDsProduto').send_keys(Description + ' - ' + PartNumber + ' ou ' + AltPartNumber)
    driver.find_element(By.ID,'btnSalvarJS').click()
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    ItemCode = driver.find_element(By.ID,'td_cPRODUTO_x_sCdProdutoEmpresa').text
    sleep(2)
    #   Removed the second tab from website  #
    driver.find_element(By.ID,'tabAba11').click()
    sleep(1)
    driver.find_element(By.ID,'check_64').click()
    driver.find_element(By.ID,'ctl00_conteudoBotoes_btnConfirmar').click()
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    print('Item Cadastrado')
    #with open ('.\\PartNumber_Log - ' + Current_Date + '.txt', 'a') as log_file:
    #    log_file.write(str(datetime.now())[0:19] + ' > ')
    #    log_file.write(PartNumber + ' - Codigo Portal: ' + ItemCode + '\n')
    # Approving item #
    sleep(0.5)

# Saving registered item
    workbook = load_workbook(filename=XL_PATH)
    sheet = workbook.active
    sheet["D"+str(i+2)] = ItemCode
    workbook.save(filename=XL_PATH)
    sleep(1)
#
    driver.find_element(By.XPATH,'//*[@id="tabAbas"]/tbody/tr/td/table[1]/tbody/tr/td[2]/a').click()
    driver.find_element(By.ID,'btnEnviarAprovacao').click() #Send to be approved
    sleep(1)
    PopUp = Alert(driver)
    PopUp.accept()
    print('Cadastro Enviado para Aprovação')
    ##################
    sleep(2)
print('Processo Finalizado')