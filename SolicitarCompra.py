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


SolicitationDesc = 'Resumo teste'#str(input('Resumo da SC: '))
SolicitationMotive = 'Justificativa teste'#str(input('Justificativa: '))
driver = webdriver.Chrome("C:\\Users\\guiga\\Documents\\PyScripts\\ExternalDrivers\\chromedriver.exe")

EXC = pd.ExcelFile("C:\\Users\\guiga\\Sistema FIEB\\Projeto OrientAgro - Documentos\\General\\Documentos de Compras\\TR1\\TR1.xlsx")
EXC = pd.ExcelFile.parse(EXC)

# ******************************** First Loggin ********************************** #
driver.get('https://compras.fieb.org.br/core/default.aspx?U=637719883608560754')
driver.find_element(By.ID,'ctl00_ctl11_tbxLogin').send_keys('guilherme.freire')
driver.find_element(By.ID,'ctl00_ctl11_tbxSenha').send_keys('meridian99')
driver.find_element(By.ID,'ctl00_ctl11_btnAcessar').click()
sleep(2)
# ******************************************************************************* #

# Creating a New Solicitation #
driver.get('https://compras.fieb.org.br/ordemcompra/OrdemCompraManutencao.aspx')
driver.find_element(By.ID, '_cORDEM_COMPRA_x_nCdTipoOrdemCompra').send_keys('Aquisição de materiais')
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
driver.find_element(By.ID,'ckbClasse_1465').click()
driver.switch_to.window(parent_window)
sleep(1)
driver.find_element(By.ID, 'btnSalvar').click()
sleep(0.5)
PopUp = Alert(driver)
PopUp.accept()
sleep(0.5)
SolicitationCode = driver.find_element(By.ID, '_cORDEM_COMPRA_x_sCdOrdemCompraEmpresa').text
print(SolicitationCode)
# Acessing the third window # (Currently Unavailible)
#driver.find_element(By.ID, 'tabAba2').click()
#driver.find_element(By.ID, 'img').click()
#parent_window = driver.current_window_handle
#child_window = driver.window_handles[1]
#driver.switch_to.window(child_window)
#driver.find_element(By.ID,'ckbClasse_17124').click()
#driver.switch_to.window(parent_window)
#sleep(1)
#driver.find_element(By.ID, '_cORDEM_COMPRA_x_nCdAcao').send_keys('OUTROS')
#driver.find_element(By.ID, '_cORDEM_COMPRA_x_nCdFontePagadora').send_keys('PROJETO')
#driver.find_element(By.ID, 'img').click()

#parent_window = driver.current_window_handle
#child_window = driver.window_handles[1]
#driver.switch_to.window(child_window)
#driver.find_element(By.ID,'ctl00_campoPesquisa_sNmUsuario').send_keys('Jovelino')
#sleep(0.5)
#driver.find_element(By.ID, 'rdbList').click()
#driver.find_element(By.ID, 'ctl00_conteudoBotoes_btnConfirmar').click()
#driver.switch_to.window(parent_window)


#driver.find_element(By.ID, 'btnSalvarJS').click()

#Incomplete
# **************** #

driver.find_element(By.ID, 'tabAba1').click()

for i in range(0, len(EXC)):
    ItemCode = int(EXC['Código Portal'][i])
    print('Inserindo Item: ' + str(ItemCode))
    # Including Individual Itens #
    driver.find_element(By.ID, 'ctl00_conteudoPagina_btnIncluir').click()
    sleep(1)
    parent_window = driver.current_window_handle
    sleep(1)
    child_window = driver.window_handles[1]
    driver.switch_to.window(child_window)
    driver.find_element(By.ID, 'ctl00_campoPesquisa_txtProduto').send_keys(ItemCode)
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
print('Solicitação Finalizada ' + SolicitationCode)

#acessar centro de custo
#Sistemas embarcados(305)
#Ação: Outros
#Fonte: Projeto
#Gestor: 
#Fiscal