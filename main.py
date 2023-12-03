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
        namaProduk = sheetRange['B' + str(i)].value
        noProduk = sheetRange['C' + str(i)].value
        unitPengukuran = sheetRange['D' + str(i)].value
        berlakuSampai = sheetRange['E' + str(i)].value
        jumlahStok = sheetRange['F' + str(i)].value
        jumlahStokInden = sheetRange['G' + str(i)].value
        kbki = sheetRange['H' + str(i)].value
        kategori = sheetRange['I' + str(i)].value


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
        ).send_keys(namaProduk)

        WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, 'inpNmr'))
        ).send_keys(noProduk)

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
        driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys(unitPengukuran)
        actions.key_down(Keys.ENTER).key_up(Keys.ENTER).perform()

        berlaku_sampai = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@id='berlaku_sampai']"))
            )
        
        berlaku_sampai.send_keys(Keys.CONTROL + "a")
        berlaku_sampai.send_keys(Keys.BACKSPACE)
        berlaku_sampai.send_keys(berlakuSampai)
        

        
        jumlah_stok_form = driver.find_element('name', 'produk.jumlah_stok_form')
        jumlah_stok_form.clear()
        jumlah_stok_form.send_keys(jumlahStok)

        
        jumlah_stok_inden_form = driver.find_element('name', 'produk.jumlah_stok_inden_form')
        jumlah_stok_inden_form.clear()
        jumlah_stok_inden_form.send_keys(jumlahStokInden)

        simpan = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='tab1']//button[@id='save']"))
        )
        simpan.click()

        time.sleep(2)

        #tab2
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

        # tab3
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
                NIB = sheetRangeSpek['G' + str(i)].value
                SBUJK = sheetRangeSpek['H' + str(i)].value
                jenisKualifikasiUsaha = sheetRangeSpek['I' + str(i)].value
                lingkupPekerjaan = sheetRangeSpek['J' + str(i)].value
                masaPemeliharaan = sheetRangeSpek['K' + str(i)].value
                gambarTeknis = sheetRangeSpek['L' + str(i)].value
                spesifikasiTeknis = sheetRangeSpek['M' + str(i)].value
                suratPerjanjianMaterial = sheetRangeSpek['N' + str(i)].value
                masaBerlakuPerjanjianMaterial = sheetRangeSpek['O' + str(i)].value
                peralatanKerjaUtama = sheetRangeSpek['P' + str(i)].value
                personilManajerial = sheetRangeSpek['Q' + str(i)].value
                pengalamanKerjaPerusahaan = sheetRangeSpek['R' + str(i)].value
                RKK = sheetRangeSpek['S' + str(i)].value
                SKP = sheetRangeSpek['T' + str(i)].value
                komponenBiaya = sheetRangeSpek['U' + str(i)].value

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

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685469']"))
                ).send_keys(NIB)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685470']"))
                ).send_keys(SBUJK)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685471']"))
                ).send_keys(jenisKualifikasiUsaha)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685472']"))
                ).send_keys(lingkupPekerjaan)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685473']"))
                ).send_keys(masaPemeliharaan)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685474']"))
                ).send_keys(gambarTeknis)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685475']"))
                ).send_keys(spesifikasiTeknis)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685476']"))
                ).send_keys(suratPerjanjianMaterial)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685477']"))
                ).send_keys(masaBerlakuPerjanjianMaterial)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685478']"))
                ).send_keys(peralatanKerjaUtama)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685479']"))
                ).send_keys(personilManajerial)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685480']"))
                ).send_keys(pengalamanKerjaPerusahaan)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685481']"))
                ).send_keys(RKK)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685482']"))
                ).send_keys(SKP)

                WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//textarea[@name='attr-id-17685483']"))
                ).send_keys(komponenBiaya)

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

        simpan3 = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@id='tab3']//button[@id='save']"))
        )
        simpan3.click()


        #tambah produk
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