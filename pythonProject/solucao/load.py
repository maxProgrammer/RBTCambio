# importa bibliotecas
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# cria objeto chrome drive
nav = webdriver.Chrome(executable_path='chromedriver')

# pesquisa cotação dolar
nav.get("https://www.google.com/")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    "cotação dólar")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacao_dolar = nav.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(f"Dólar hoje: R$ %.2f" % float(cotacao_dolar))

# pesquisa cotação euro
nav.get("https://www.google.com/")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(
    "cotação euro")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
nav.implicitly_wait(4)
cotacao_euro = nav.find_element_by_xpath(
    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(f"Euro Hoje: R$ %.2f" % float(cotacao_euro))

# pesquisa cotação ouro
nav.get("https://www.melhorcambio.com/ouro-hoje")
aba_original = nav.window_handles[0]
cotacao_ouro = nav.find_element_by_xpath('//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(',', '.')
print(f"Ouro hoje: R$ %.2f" % float(cotacao_ouro))

# fecha a janela após finalizar consultas
nav.quit()

import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")

# atualizar a contação
# atualizar coluna cotação aonde coluna moeda for igual a Dólar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = round(float(cotacao_dolar), 2)

# atualizar coluna cotação aonde coluna moeda for igual a Euro
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = round(float(cotacao_euro), 2)

# atualizar coluna cotação aonde coluna moeda for igual a Ouro
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = round(float(cotacao_ouro), 2)

print(tabela)

# atualizar o preço de compra => preço original * cotação
tabela["Preço Base Reais"] = round(tabela["Preço Base Original"] * tabela["Cotação"], 2)

# atualizar o preço de venda => preço de compra * margem
tabela["Preço Final"] = round(tabela["Preço Base Reais"] * tabela["Margem"], 2)

print(tabela)
