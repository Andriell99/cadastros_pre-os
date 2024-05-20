from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Inicializa o driver do Selenium
driver = webdriver.Chrome()
driver.get('https://www.novaliderinformatica.com.br/computadores')

# Encontrando os nomes dos itens
titulos = driver.find_elements(By.XPATH, "//a[@class='nome-produto']")
# Encontrando os preços promocionais dos itens
precos = driver.find_elements(By.XPATH, "//strong[@class='preco-promocional']")

# Definindo nome da planilha
planilha = openpyxl.Workbook()
# Removendo a planilha padrão
sheet_produtos = planilha.active
sheet_produtos.title = 'produtos'
# Selecionando a página de produtos e as colunas a serem alimentadas
sheet_produtos['A1'].value = 'produto'
sheet_produtos['B1'].value = 'preço'

# Loop de iteração para inserção de dados na planilha
for titulo, preco in zip(titulos, precos):
    sheet_produtos.append([titulo.text, preco.text])

# Salvando a planilha
planilha.save('produtos.xlsx')

# Fechando o driver
driver.quit()
