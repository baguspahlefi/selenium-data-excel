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

driver.get("https://e-katalog.lkpp.go.id/v3/katalog/produk/create/76086456") 
actions = ActionChains(driver)
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
    kbki = sheetRange['F' + str(i)].value


    driver.find_element('id','namaProduk').send_keys(name)
    driver.find_element('id','inpNmr').send_keys(inpNmr)
    time.sleep(2)
    

    WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, 'select2-selectDiproduksiDiIndonesia-container'))
    ).click()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, 'select2-selectTenagaKerjaIndonesia-container'))
    ).click()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, 'select2-selectBahanBakuDalamNegeri-container'))
    ).click()
    time.sleep(2)
    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
    actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

    actions.key_down(Keys.ARROW_DOWN).key_up(Keys.ARROW_DOWN).perform()

    # Tekan tombol Enter untuk memilih opsi yang sedang di-highlight
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

    WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//a[contains(@href,"#tab2")]'))
    ).click()

    search = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//input[contains(@type,"search")]'))
    )
    search.click()
    search.send_keys(kbki)

    time.sleep(2)


    # Constructing the XPath with the variable
    xpath = f'//button[contains(@kbkiid,"{kbki}")]'

    # Using WebDriverWait to wait for the element to be clickable and then clicking it
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    ).click()

    time.sleep(5)
    

    

    time.sleep(2)
    i = i+1

    input("Press Enter to close the browser...")

    # Tutup peramban setelah selesai
    driver.quit()