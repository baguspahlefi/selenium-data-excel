from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import time

wb = load_workbook(filename="/Users/macbook/Project/input-excel/data-ikan.xlsx")

sheetRange = wb['Sheet1']

driver = webdriver.Firefox()

driver.get("https://e-katalog.lkpp.go.id/v3/katalog/produk/create/76086456") 
driver.maximize_window()
driver.implicitly_wait(10)

#   
userName = 'CVMITRADINASTI1'
password = 'Putra@123123'

driver.find_element('name','username').send_keys(userName)
driver.find_element('name','password').send_keys(password)
driver.find_element('id','btnLoginPenyedia').click()

WebDriverWait(driver, 10).until(lambda x: x.find_element('id', "namaProduk"))

i = 2

while i <= len(sheetRange['A']):
    name = sheetRange['A' + str(i)].value
    inpNmr = sheetRange['A' + str(i)].value


    driver.find_element('id','namaProduk').send_keys(name)
    driver.find_element('id','inpNmr').send_keys(inpNmr)
    

    # Temukan elemen dropdown
# Wait for the dropdown element to be clickable
    dropdown_element = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, 'selectDiproduksiDiIndonesia'))
    )
    

    dropdown_element.click()
   
    dropdown_element.send_keys(Keys.ARROW_DOWN)
    time.sleep(3)

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    dropdown_element.send_keys(Keys.ENTER)

    time.sleep(10)


    time.sleep(2)
    i = i+1

    input("Press Enter to close the browser...")

    # Tutup peramban setelah selesai
    driver.quit()