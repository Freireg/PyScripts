import logging
import openpyxl
import pandas as pd
from time import sleep
from pandas import ExcelFile
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common import by
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import SessionNotCreatedException
import getpass


VERSION = '1.2.2'


print('______________________________________________________')
print("Automatização de Solicitação de Compras do Portal de Compras FIEB")
print('Versão ' + VERSION)
print("Desenvolvido por Guilherme Freire")
print('______________________________________________________\n\n')



user = input("Usuário: ")
passw = getpass.getpass('Senha: ')
projectType = int(input("Projeto Embrapi?\n1 -> Sim\n2 -> Não\nFavor indicar: "))

if(projectType == 1):
    pType = "Embrapii"
elif(projectType == 2):
    pType = "Aquisição de materiais"
else:
    print("Argumento inválido")

try:
    driver = webdriver.Chrome(".\\chromedriver.exe")
except SessionNotCreatedException:
    print('\n\nVersão do driver incompatível!')
    print('Necessário realizar o download do driver com versão adequada ao seu navegador')
    input('[Pressione ENTER para encerrar]')
    exit()
EXC = pd.ExcelFile(".\\Cadastro.xlsx")
EXC = pd.ExcelFile.parse(EXC)
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(options=options)

SolicitationDesc = str(EXC['Resumo'][0])#str(input('Resumo da SC: '))
SolicitationMotive = str(EXC['Justificativa'][0])#str(input('Justificativa: '))

# ******************************** First Loggin ********************************** #
driver.get('https://compras.fieb.org.br/core/default.aspx?U=637719883608560754')
driver.find_element(By.ID,'ctl00_ctl11_tbxLogin').send_keys(user)
driver.find_element(By.ID,'ctl00_ctl11_tbxSenha').send_keys(passw)
driver.find_element(By.ID,'ctl00_ctl11_btnAcessar').click()
sleep(2)
# ******************************************************************************* #

# Creating a New Solicitation #
driver.get('https://compras.fieb.org.br/ordemcompra/OrdemCompraManutencao.aspx')
driver.find_element(By.ID, '_cORDEM_COMPRA_x_nCdTipoOrdemCompra').send_keys(pType)
driver.find_element(By.ID, '_cORDEM_COMPRA_x_sDsOrdemCompra').send_keys(SolicitationDesc)
driver.find_element(By.ID, '_cORDEM_COMPRA_x_sDsJustificativa').send_keys(SolicitationMotive)
driver.find_element(By.ID, '_cORDEM_COMPRA_x_nCdAplicacao').send_keys('Orçamento')
sleep(2)
driver.find_element(By.XPATH, '//*[@id="td_cORDEM_COMPRA_x_nCdDepartamento"]/table/tbody/tr/td[2]/a').click()
sleep(1)
parent_window = driver.current_window_handle
sleep(1)
child_window = driver.window_handles[1]
driver.switch_to.window(child_window)
driver.find_element(By.ID,'ckbClasse_1433').click()
driver.switch_to.window(parent_window)
sleep(1)
driver.find_element(By.ID, 'btnSalvar').click()
sleep(1)
PopUp = Alert(driver)
PopUp.accept()
sleep(0.5)
SolicitationCode = driver.find_element(By.ID, '_cORDEM_COMPRA_x_sCdOrdemCompraEmpresa').text
print(SolicitationCode)

driver.find_element(By.ID, 'tabAba1').click()

for i in range(0, len(EXC)):
    ItemCode = str(EXC['Description'][i])

    print('Inserindo Item: ' + str(ItemCode))
    # Including Individual Itens #
    driver.find_element(By.ID, 'ctl00_conteudoPagina_btnIncluir').click()
    sleep(1)
    parent_window = driver.current_window_handle
    sleep(1)
    child_window = driver.window_handles[1]
    driver.switch_to.window(child_window)
    driver.find_element(By.ID, 'ctl00_campoPesquisa_sDsProduto').send_keys(ItemCode)
    driver.find_element(By.ID, 'ctl00_btnPesquisar').click()
    sleep(1)
    driver.find_element(By.ID, 'ckbList').click()
    driver.find_element(By.ID, 'ctl00_conteudoBotoes_btnConfirmar').click()
    sleep(0.5)
    driver.switch_to.window(parent_window)
    sleep(3)
    print('Item inserido na SC\n')

    # ************************** #
    sleep(0.5) #Debug wait


idx = 2
for i in range(0, len(EXC)):

    PartNumber = (EXC['Part Number'][i])
    Qnty = int(EXC['Quantidade'][i])

    
    sleep(1)
    driver.find_element(By.ID, 'tbxDsOrdemCompra').clear()
    driver.find_element(By.ID, 'tbxDsOrdemCompra').send_keys(PartNumber)
    driver.find_element(By.ID, 'ctl00_conteudoPagina_btPesquisar').click()
    sleep(2)
    for s in range(0, Qnty):
        driver.find_element(By.XPATH, '//*[@id="tblPesquisa"]/tbody/tr['+str(idx)+']/td[1]/span[2]/span/span/span[1]/span').click()
        
    sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="tblPesquisa"]/tbody/tr['+str(idx)+']/td[3]/span[2]/span/span/span').click()
    sleep(1)  
    driver.find_element(By.CSS_SELECTOR, 'td[class="k-today k-state-focused"]').click()              

    idx = int(idx)
    idx = idx + 3

    
driver.find_element(By.ID, "ctl00_conteudoBotoes_btSalvar").click()

print('Solicitação Finalizada ' + SolicitationCode)