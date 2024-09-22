import time
import os
import pyodbc
from bs4 import BeautifulSoup
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC

# Configuração do WebDriver
service = ChromeService(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Loga no sistema
driver.get('https://exemplo.com/login')

# Realiza login no sistema
username = driver.find_element(By.XPATH, '//*[@id="usuario"]')
password = driver.find_element(By.XPATH, '//*[@id="senha"]')
username.send_keys('seu_usuario')
password.send_keys('sua_senha')
password.send_keys(Keys.RETURN)

# Fecha pop-ups, se presentes
try:
    close_button_popup1 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="fechar-popup1"]'))
    )
    close_button_popup1.click()
except:
    pass

try:
    close_button_popup2 = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="fechar-popup2"]'))
    )
    close_button_popup2.click()
except:
    pass

# Aguarda a página carregar
WebDriverWait(driver, 20).until(EC.url_contains('Home'))

# Navega para a página home
driver.get('https://exemplo.com/home')

# Interage com a página 
try:
    filter_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id=""]'))
    )
    filter_button.click()

    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id=""]'))
    )
    select = Select(select_element)
    select.select_by_visible_text('Este Mês')

    search_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id=""]'))
    )
    search_button.click()
except:
    pass

# Aguarda o download
time.sleep(20)

# Fecha o navegador
driver.quit()

# Caminho da pasta de downloads
downloads_path = str(Path.home() / "Downloads")
files = [f for f in os.listdir(downloads_path) if f.endswith('.xls')]

# Processa o arquivo baixado
if files:
    latest_file = max([os.path.join(downloads_path, f) for f in files], key=os.path.getctime)

    if os.path.getsize(latest_file) > 0:
        try:
            with open(latest_file, 'r', encoding='utf-8') as file:
                soup = BeautifulSoup(file, 'lxml')

            table = soup.find('table')

            if table:
                rows = table.find_all('tr')

                # Conecta a um banco de dados 
                connection_string = (
                    "Driver={ODBC Driver 17 for SQL Server};"
                    "Server=servidor;"
                    "Database=base_de_dados;"
                    "UID=usuario;"
                    "PWD=senha;"
                )
                conn = pyodbc.connect(connection_string)
                cursor = conn.cursor()

                for row in rows:
                    cols = row.find_all('td')
                    if len(cols) == 5:
                        ex1, ex2, ex3, ex4, ex5 = [col.get_text(strip=True) for col in cols]
                        cursor.execute("""
                            INSERT INTO TABELA (ex1, ex2, ex3, ex4, ex5) 
                            VALUES (?, ?, ?, ?, ?)
                        """, ex1, ex2, ex3, ex4, ex5)

                conn.commit()
                cursor.close()
                conn.close()
        except:
            pass
