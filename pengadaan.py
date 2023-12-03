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
sheetRangeSpek = wb['Sheet2']

driver = webdriver.Firefox()

driver.get("https://e-katalog.lkpp.go.id/v3/katalog/produk/create/76105689") 
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
            search_kategori = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, " //input[@role='textbox']"))
            )
            search_kategori.send_keys('Mistar Muka Air')
            actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)

            WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685649']"))
            ).send_keys('asu')

        elif kategori == "Struktur Baja" :
            search_kategori = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, " //input[@role='textbox']"))
            )
            search_kategori.send_keys('Struktur Baja')
            actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)
            j = 2
            while j <= len(sheetRangeSpek['B']):
                keteranganLainya = sheetRangeSpek['B' + str(i)].value
                TKDN = sheetRangeSpek['C' + str(i)].value
                perusahaan = sheetRangeSpek['D' + str(i)].value
                lokasi = sheetRangeSpek['E' + str(i)].value
                KBLI = sheetRangeSpek['F' + str(i)].value

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685464']"))
                ).send_keys(keteranganLainya)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685465']"))
                ).send_keys(TKDN)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685466']"))
                ).send_keys(perusahaan) 

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685467']")) 
                ).send_keys(lokasi) 

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685468']"))
                ).send_keys(KBLI)

                j = i+1

           

        elif kategori == "Pengadaan Barang Peralatan/Mesin/Perlengkapan Budidaya" :
            search_kategori = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, " //input[@role='textbox']"))
            )
            search_kategori.send_keys('Pengadaan Barang')
            actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)

            WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685037']"))
            ).send_keys('pengadaan mesin')

        elif kategori == "Pengadaan Barang Sarana dan Prasana Laboratorium Perikanan Budidaya" :
            search_kategori = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, " //input[@role='textbox']"))
            )
            search_kategori.send_keys('Pengadaan Barang')
            actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
            actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
            time.sleep(2)

            WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685076']"))
            ).send_keys('Pengadaan Barang Sarana dan Prasana Laboratorium Perikanan Budidaya')

        

        
        

        


        

        

        time.sleep(20)
        i = i+1
finally:
    # Tutup peramban setelah selesai
    driver.quit()