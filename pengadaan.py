from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import load_workbook
import time

wb = load_workbook(filename="data-ikan.xlsx")

sheetRange = wb['Sheet1']

driver = webdriver.Firefox()

driver.get("https://e-katalog.lkpp.go.id/v3/katalog/produk/create/76103719") 
actions = ActionChains(driver)
driver.maximize_window()
driver.implicitly_wait(10)

#   
userName = 'CVMITRADINASTI1'
password = 'Putra@123123'

driver.find_element('name','username').send_keys(userName)
driver.find_element('name','password').send_keys(password)
driver.find_element('id','btnLoginPenyedia').click()

time.sleep(5)


i = 2


try:
    while i <= len(sheetRange['A']):
        judulEtalase = sheetRange['A' + str(i)].value
        kategori = sheetRange['G' + str(i)].value

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//a[contains(@href,"#tab3")]'))
        ).click()

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-level1-container']"))
        ).click()

        if kategori == "Mistar Muka Air":
            
            WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685649']"))
            ).send_keys('asu')


        

        

        time.sleep(20)
        i = i+1
finally:
    # Tutup peramban setelah selesai
    driver.quit()