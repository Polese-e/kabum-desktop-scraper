import openpyxl.workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

# Access the website https://www.kabum.com.br/computadores/pc
driver = webdriver.Chrome()
driver.get('https://www.kabum.com.br/computadores/pc')

# Extract all titles
titles = driver.find_elements(By.XPATH,'//span[@class="sc-d79c9c3f-0 nlmfp sc-9d1f1537-16 fQnige nameCard"]')
# Extract all prices
prices = driver.find_elements(By.XPATH,'//span[@class="sc-b1f5eb03-2 iaiQNF priceCard"]')

# Creating the workbook
workbook = openpyxl.Workbook()
# Creating the 'products' sheet
workbook.create_sheet('products')
# Selecting the products sheet
sheet_products = workbook['products']
sheet_products['A1'].value = 'Product'
sheet_products['B1'].value = 'Price'

# Inserting titles and prices into the spreadsheet
for title, price in zip(titles, prices):
    sheet_products.append([title.text, price.text])

workbook.save('products.xlsx')
