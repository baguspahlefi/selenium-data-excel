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

driver.get("https://e-katalog.lkpp.go.id/v3/katalog/produk/create") 
actions = ActionChains(driver)
driver.maximize_window()
driver.implicitly_wait(10)

#   
userName = 'CVMITRADINASTI1'
password = 'Putra@123123'

driver.find_element('name','username').send_keys(userName)
driver.find_element('name','password').send_keys(password)
driver.find_element('id','btnLoginPenyedia').click()


i = 2


try:
    while i <= len(sheetRange['A']):
        judulEtalase = sheetRange['A' + str(i)].value
        name = sheetRange['B' + str(i)].value
        inpNmr = sheetRange['B' + str(i)].value
        kbki = sheetRange['F' + str(i)].value


        etalase_produk = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'select2-komoditas-container'))
        )
        etalase_produk.click()
        driver.find_element('xpath','//input[contains(@type,"search")]').send_keys(judulEtalase)
        actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        time.sleep(5)

        pengumuman = driver.find_element('id','select2-usulan-container')
        webdriver.ActionChains(driver).move_to_element(pengumuman).click(pengumuman).perform()
        driver.find_element('xpath','//input[contains(@type,"search")]').send_keys(judulEtalase)
        actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        driver.find_element('id','save').click()


        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'namaProduk'))
        ).send_keys(name)
        driver.find_element('id','inpNmr').send_keys(inpNmr)
        time.sleep(2)

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'select2-selectDiproduksiDiIndonesia-container'))
        ).click()
        actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'select2-selectTenagaKerjaIndonesia-container'))
        ).click()
        actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()


        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'select2-selectBahanBakuDalamNegeri-container'))
        ).click()
        actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'unit_pengukuran'))
        ).click()

        driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys('unit')
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        berlaku_sampai = driver.find_element('id', 'berlaku_sampai')
        berlaku_sampai.clear()
        berlaku_sampai.send_keys("31-12-2023")

        
        jumlah_stok_form = driver.find_element('name', 'produk.jumlah_stok_form')
        jumlah_stok_form.clear()
        jumlah_stok_form.send_keys("1000")

        
        jumlah_stok_inden_form = driver.find_element('name', 'produk.jumlah_stok_inden_form')
        jumlah_stok_inden_form.clear()
        jumlah_stok_inden_form.send_keys("100")

        simpan = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='tab1']//button[@id='save']"))
        )
        simpan.click()

        time.sleep(2)

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//a[contains(@href,"#tab2")]'))
        ).click()

        search = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//input[contains(@type,"search")]'))
        )
        search.click()
        search.send_keys(kbki)

        # Constructing the XPath with the variable
        xpath = f'//button[contains(@kbkiid,"{kbki}")]'
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        ).click()

        
        simpan2 = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='tab2']//button[@id='save']"))
        )
        simpan2.click()

        toogle_produk = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='dropdown-toggle' and normalize-space()='Produk']"))
        )
        toogle_produk.click()
        
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@class='row']//a[contains(text(),'Tambah Produk')]"))
        ).click()


        

        

        time.sleep(2)
        i = i+1
finally:
    # Tutup peramban setelah selesai
    driver.quit()