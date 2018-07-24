# papulic 2018
# encoding=utf8
import os
import shutil
import openpyxl
import time
import sys
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
reload(sys)
sys.setdefaultencoding('utf8')

# #####################################################################################################################
# #####################################################################################################################
# #####################################################################################################################

# cuvanje svih rezultata (stanje u svim magacinima i ostali podaci se cuvaju u tekstualnim fajlovima)
CUVAJ_SVE = True

# pauza je u sekundama!!!
PAUZA_IZMEDJU_PRETRAGA = 30

# korisnik za gazelu
korisnik_gazela = "fiatino"
lozinka_gazela = "fiatino123"

# #####################################################################################################################
# #####################################################################################################################
# #####################################################################################################################

param_error = """
Pogresan broj parametara..
Kako koristiti skriptu: rv_trade.py <textualni fajl sa spiskom artikala za pretragu> <sajt za pretragu>
Moguci sajtovi: 'gazela'
"""

MOGUCI_SAJTOVI = ['gazela']

# sajtovi
gazela = False
# Artikli za pretragu
artikli_za_pretragu = []

# Provera ispravnosti parametara
if not len(sys.argv) == 3:
    sys.exit(param_error)
else:
    try:
        lista = open(sys.argv[1]).readlines()
    except IOError:
        sys.exit("Ne postoji fajl: '{0}'!".format(sys.argv[1]))
    for artikal in lista:
        if artikal != "\n":
            artikal = artikal.replace("\n", "")
            if artikal not in artikli_za_pretragu:
                artikli_za_pretragu.append(artikal)
    sajt = sys.argv[2]
    if sajt not in MOGUCI_SAJTOVI:
        sys.exit("Pogresan sajt: '{0}'!".format(sys.argv[2]))
    else:
        if sajt == "gazela":
            gazela = True

# excel
XLSX_FILE_TEMP = "rezultati_temp.xlsx"
XLSX_FILE = "rezultati.xlsx"

# ____________________________________________________________________________________________________
if os.path.exists(XLSX_FILE_TEMP):
    book = openpyxl.load_workbook(XLSX_FILE_TEMP)
else:
    wb = openpyxl.Workbook()
    for i in MOGUCI_SAJTOVI:
        wb.create_sheet(i)
    wb.remove_sheet(wb.get_sheet_by_name('Sheet'))
    wb.save(XLSX_FILE_TEMP)
    book = openpyxl.load_workbook(XLSX_FILE_TEMP)



# get the path of ChromeDriverServer
dir = os.path.dirname(__file__)
chrome_driver_path = "chromedriver.exe"

# create a new Chrome session
driver = webdriver.Chrome(chrome_driver_path)
driver.implicitly_wait(30)
# mora full screen , u suprotnom nece naci pretragu !!
driver.maximize_window()

if gazela:
    driver.get("http://gazela.co.rs/")
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPass")

    username.send_keys(korisnik_gazela)
    password.send_keys(lozinka_gazela)

    driver.find_element_by_id("Button1").click()
    for artikal in artikli_za_pretragu:
        print "Pretrazujem artikal '{0}'...\n".format(artikal)
        time.sleep(2)

        # pretraga
        nasao_pretragu = False
        while not nasao_pretragu:
            try:
                driver.find_element_by_xpath('//a[@href="TopMotive.aspx"]').click()
                nasao_pretragu = True
            except:
                pass

        time.sleep(2)

        driver.switch_to.frame(driver.find_element_by_id("Main_iframe1"))

        searchbox = driver.find_element_by_id("home_txt_art_direkt")

        # unesi artikal za pretragu
        searchbox.send_keys(artikal)
        time.sleep(1)

        driver.find_element_by_id("home_imgBtn_art_direkt").click()

        time.sleep(3)
        html = driver.page_source
        # ucitavanje cele stranice
        soup = BeautifulSoup(html, "lxml")
        uspesna_pretraga = True

        tabela = soup.find('table', id="main_artikel_panel_maintable")

        brendovi = []
        for row in tabela.find_all('tr'):
            if row.has_attr('class'):
                if "main_artikel_panel_tr_einspeiser" in row.attrs['class']:
                    columns = row.find_all('td')
                    for col in columns:
                        if col.has_attr('colspan'):
                            brend = col.find('span').text
                            brendovi.append([brend])
                elif "main_artikel_panel_tr_artikel" in row.attrs['class']:
                    if "artikel1" in row.attrs['row_type']:
                        brendovi[-1].append({})
                        status_i_cena = row.find_all('td')
                        for col in status_i_cena:
                            if col.has_attr('class'):
                                if 'tc_number' in col.attrs['class']:
                                    sifre = col.find_all('a')
                                    for a in sifre:
                                        if a.has_attr('title'):
                                            if a.attrs['title'].startswith("Sifra dobavljacevog artikla"):
                                                sifra_dobavljacevog_artikla = a.contents[0].text
                                                brendovi[-1][-1]["sifra dobavljacevog artikla"] = sifra_dobavljacevog_artikla
                                            elif a.attrs['title'].startswith("Kataloski broj"):
                                                kataloski_broj_proizvodjaca = a.contents[0].text
                                                brendovi[-1][-1]["kataloski broj proizvodjaca"] = kataloski_broj_proizvodjaca

        cene = driver.find_elements_by_tag_name('img')
        # sve_cene = []
        for cena in cene:
            title = cena.get_attribute('title')
            if title.startswith('Pitanje o trenutnim kolicinama/ceni'):
                # sve_cene.append(cena)
                driver.execute_script("arguments[0].click();", cena)
                time.sleep(2)
                driver.switch_to.frame(driver.find_elements_by_class_name("cboxIframe")[0])
                time.sleep(1)
                master = driver.find_elements_by_id("masterpane")[0]
                html_text = master.text
                html_cena = master.text.split('\n')
                if len(html_cena) > 11:
                    artikal_pretraga = html_cena[0]
                    stanje = html_cena[3]
                    if stanje.startswith("Dostupno"):
                        stanje = "IMA"
                    elif stanje.startswith("Nije dostupno"):
                        stanje = "NEMA"
                    else:
                        stanje = "NEPOZNATO"
                    vp_cena = html_cena[11]
                    for brend in brendovi:
                        for num, a in enumerate(brend):
                            if num > 0:
                                if a['sifra dobavljacevog artikla'] == artikal_pretraga:
                                    a["vp cena"] = vp_cena
                                    a["na stanju"] = stanje
                                    print artikal_pretraga
                                    if CUVAJ_SVE:
                                        if not os.path.exists("rezultati"):
                                            os.makedirs("rezultati")
                                        ime_fajla = artikal + "-" + brend[0] + "-" + a["kataloski broj proizvodjaca"]
                                        for ch in ['&', '#', '/', '\\', '<', '>', '"', '?', "'", ':', ';', '|', '*', '!', ' ']:
                                            if ch in ime_fajla:
                                                ime_fajla = ime_fajla.replace(ch, "")
                                        temp = open("rezultati\\" + ime_fajla + ".txt", 'w')
                                        temp.write(html_text)
                                        temp.close()

                driver.switch_to.default_content()
                # driver.implicitly_wait(30)
                time.sleep(1)
                driver.switch_to.frame(driver.find_element_by_id("Main_iframe1"))
                time.sleep(1)
                element = driver.find_element_by_id('cboxClose')
                driver.execute_script("arguments[0].click();", element)
                time.sleep(1)

        worksheet = book['gazela']
        row_idx = worksheet.max_row + 2
        worksheet.cell(row=row_idx, column=1).value = artikal
        worksheet.cell(row=row_idx + 1, column=2).value = "HEAD1"
        worksheet.cell(row=row_idx + 1, column=3).value = "HEAD2"
        worksheet.cell(row=row_idx + 1, column=4).value = "HEAD3"
        worksheet.cell(row=row_idx + 1, column=5).value = "HEAD4"
        row_idx += 2
        for brend in brendovi:
            for num, i in enumerate(brend):
                if num > 0:
                    worksheet.cell(row=row_idx, column=2).value = brend[0]
                    if 'kataloski broj proizvodjaca' in i:
                        worksheet.cell(row=row_idx, column=3).value = i['kataloski broj proizvodjaca']
                    else:
                        worksheet.cell(row=row_idx, column=3).value = '---'
                    if 'vp cena' in i:
                        vp_cena = i['vp cena'].split('VP Cena ')[1].split(' RSD')[0]
                        if vp_cena[0].isdigit():
                            vp_cena = float(vp_cena)
                        worksheet.cell(row=row_idx, column=4).value = vp_cena
                    else:
                        worksheet.cell(row=row_idx, column=4).value = '---'
                    if 'na stanju' in i:
                        worksheet.cell(row=row_idx, column=5).value = i['na stanju']
                    else:
                        worksheet.cell(row=row_idx, column=5).value = '---'
                    row_idx += 1

        book.save(XLSX_FILE_TEMP)
        if os.path.exists(XLSX_FILE):
            os.remove(XLSX_FILE)
        shutil.copyfile(XLSX_FILE_TEMP, XLSX_FILE)
        driver.switch_to.default_content()
        if artikal != artikli_za_pretragu[-1]:
            print "\nCekam {sekunde} sekundi".format(sekunde=str(PAUZA_IZMEDJU_PRETRAGA))
            time.sleep(PAUZA_IZMEDJU_PRETRAGA)
    driver.quit()
    print "\nUspesna pretraga!"
