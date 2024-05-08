from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

driver = webdriver.Chrome()
driver.get("https://www.novaliderinformatica.com.br/computadores-gamers")

nome_produto = driver.find_elements(By.XPATH, "//a[@class='nome-produto']")
preco_produto = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

open = openpyxl.Workbook()
open.create_sheet('produtos')
sheet_produtos = open['produtos']
sheet_produtos['A1'].value = 'Produtos'
sheet_produtos['B1'].value = 'Preço'

for titulos, precos in zip(nome_produto, preco_produto):
    try:
        sheet_produtos.append([titulos.text, precos.text])
    except Exception as e:
        print(e)
open.save('produtos.xlsx')