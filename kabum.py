from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time
import os
import unicodedata

print('Pasta atual:', os.getcwd())

def limpa_texto(texto):
    texto = texto.lower()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto)
                    if unicodedata.category(c) != 'Mn')
    return texto

service = Service(r'C:\Users\luisa\Downloads\edgedriver_win64\msedgedriver.exe')
driver = webdriver.Edge(service=service)
driver.maximize_window()

driver.get('https://www.kabum.com.br/')

wait = WebDriverWait(driver, 15)

busca = wait.until(EC.element_to_be_clickable((By.NAME, 'query')))
busca.click()
busca.clear()

termo = 'placa de vídeo'
for letra in termo:
    busca.send_keys(letra)
    time.sleep(0.1)

time.sleep(0.5)
busca.send_keys(Keys.ENTER)

wait.until(EC.url_contains('busca'))

wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'placa')]")))

time.sleep(3)  

total_scrolls = 5
for _ in range(total_scrolls):
    driver.execute_script("window.scrollBy(0, 1000)")
    time.sleep(2)

wb = Workbook()
ws = wb.active
ws.append(['Nome do Produto', 'Preço'])

produtos = driver.find_elements(By.CLASS_NAME, 'productCard')

nomes_vistos = set()
contador = 0
limite = 5

for produto in produtos:
    try:
        nome = produto.find_element(By.CLASS_NAME, 'nameCard').text
        preco = produto.find_element(By.CLASS_NAME, 'priceCard').text
        nome_limpo = limpa_texto(nome)
        if 'placa de video' in nome_limpo and nome_limpo not in nomes_vistos:
            ws.append([nome, preco])
            nomes_vistos.add(nome_limpo)
            contador += 1
            if contador >= limite:
                break
    except Exception as e:
        print(f'Erro ao capturar produto: {e}')

arquivo = 'produtos_kabum.xlsx'
wb.save(arquivo)

print(f'Extração concluída e planilha salva em: {os.path.abspath(arquivo)}')

driver.quit()
