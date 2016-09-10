# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import contextlib
from contextlib import contextmanager
import datetime
import xlrd
import xlwt
import os
import errno
import arial10
import lxml
 

        
now = datetime.datetime.now()
today = now.strftime('%Y')+"-"+now.strftime('%m')+"-"+now.strftime('%d')
url = "http://www.dipsco.unitn.it/Logistica/Vedi/lezioniEventi/agenda_periodoc.jsp?menuSxIndex=4"

rows = []
flagged_a = []
flagged_m = []


# APRE IL FILE .XLS CONTENENTE LA LISTA DELLE AUDIOLEZIONI
def grab_audio_list():
    workbook = xlrd.open_workbook('audio_lezioni_semestre.xls')
    worksheet = workbook.sheet_by_name('Sheet1')

    lezioni_xls = [(worksheet.row_values(i)[0]).lower().encode('ascii', 'ignore') for i in range(worksheet.nrows)]

    return lezioni_xls

    
# PRENDE LA TABELLA DELLE LEZIONI DEL GIORNO
def selenium_form():
    phantomjs_path = "phantomjs.exe"

    @contextmanager
    def quitting(quitter):
        try:
            yield quitter
        finally:
            quitter.quit()
        
    with quitting(webdriver.PhantomJS(executable_path=phantomjs_path,
                                                service_log_path=os.path.devnull)) as driver:
        driver.set_window_size(1400, 1000)                     
        driver.get(url)
        # CAMBIARE LA DATA PER TESTING
        # driver.execute_script('document.getElementById("giorno_inizio").removeAttribute("readonly")')
        # driver.execute_script('document.getElementById("giorno_fine").removeAttribute("readonly")')
        # date_field_1 = driver.find_element_by_name('giorno_inizio')
        # date_field_1.clear()
        # date_field_1.send_keys(today)
        # date_field_1.send_keys("2015-04-27")
        # date_field_2 = driver.find_element_by_name('giorno_fine')
        # date_field_2.clear()
        # date_field_2.send_keys(today)
        # date_field_2.send_keys("2015-04-27")

        submit = driver.find_element_by_xpath("//input[@value='Continue']")
        submit.click()
        # SI CARICA LA PAGINA CON LA TABELLA DELLE AUDIOLEZIONI, DELAY, TRY PER ASPETTARLA
        delay = 2  # seconds
        try:
            page = WebDriverWait(driver, delay).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="str-mainbox"]/div[4]/fieldset/table'))
            )
            # print "Table is ready!"
            return page.get_attribute('innerHTML')
        except TimeoutException:
            print "Loading took too much time!"


# DALLA TABELLA FORMATO HTML SPUTA IL CONTENUTO DELLE RIGHE            
def table_scrap(table_html):
    lezioni_oggi = []
    table_html = table_html.encode('utf-8', 'ignore')
    table = BeautifulSoup(table_html, "lxml")  
    
    # ROWS GLOBALE E' UTILE ANCHE ALLA CREAZIONE DELLA TABELLA FINALE
    global rows
    
    rows = table.find_all('tr')
    # print "rows"
    # print rows
    for row in rows:
        row_text = row.get_text(" ", strip=True).lower()
        lezioni_oggi.append(row_text.encode('ascii', 'ignore'))
    # LEZIONI OGGI E' UNA LISTA DI STRINGHE, CIASCUNA CONTENENTE IL TESTO DI OGNI RIGA
    # print lezioni_oggi
    return lezioni_oggi


# COMPARA LE LEZIONI DEL GIORNO ALLE AUDIOLEZIONI E SPUTA UNA LISTA DI INDEX FLAGGATI
def compare(lezioni_oggi, lezioni_xls):
    global flagged_m
    global flagged_a
    for idx, val in enumerate(lezioni_oggi):
        if any(lez in val for lez in lezioni_xls):
            flagged_a.append(idx)
            # print "flagged_a"
            # print idx
        n_aule = val.count("aula")
        if n_aule > 1:
            flagged_m.append(idx)
            # print "flagged_m"
            # print idx


# AUTORESIZING DELLE COLONNE NELL'XLS DI OUTPUT CON ARIAL10            
class FitSheetWrapper(object):
    def __init__(self, sheet):
        self.sheet = sheet
        self.widths = dict()

    def write(self, r, c, label='', *args, **kwargs):
        self.sheet.write(r, c, label, *args, **kwargs)
        width = int(arial10.fitwidth(label))
        if width > self.widths.get(c, 0):
            self.widths[c] = width
            self.sheet.col(c).width = width

    def __getattr__(self, attr):
        return getattr(self.sheet, attr)


# CREAZIONE DEL FILE DI OUTPUT FINALE        
def colour_table():
    r = 0
    style = xlwt.easyxf('font: bold 1;')
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = FitSheetWrapper(workbook.add_sheet(today, cell_overwrite_ok=True))
    # sheet.write(r, 0, "".join("TITLE"))
    for idx, row in enumerate(rows):
        c = 0
        r += 1
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        for n_col, col in enumerate(cols):
            if idx in flagged_a or idx in flagged_m:
                sheet.write(r, c, " ".join(col.split()), style)
            else:
                sheet.write(r, c, " ".join(col.split()))
            c += 1

            if n_col == 3:
                if idx in flagged_m:
                    sheet.write(r, c, "Aule collegate", style)
                    # print idx, "  flagged_m"
                if idx in flagged_a:
                    sheet.write(r, c, "Audiolezione", style)
                    # print idx, "  flagged_a"
         
    silent_remove(today + '.xls')
    workbook.save(today + '.xls')
    # print "Done."


# MODULO PER NON OVERLAPPARE I FILE DI OUTPUT
def silent_remove(filename):
    try:
        os.remove(filename)
    except OSError as e:
        if e.errno != errno.ENOENT:  # errno.ENOENT = no such file or directory
            raise  # re-raise exception if a different error occured

compare(table_scrap(selenium_form()), grab_audio_list())
colour_table()