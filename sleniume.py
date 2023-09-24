import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import pandas as pd
import os

try:
    
    # =============================================================================
    # Disable chrome items
    # =============================================================================
    # end_pro.set("")
    # profile = "C:\\Users\\Faizan\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 25"
    # uc.TARGET_VERSION = 116 
    options = uc.ChromeOptions()
    options.add_argument('--disable-notifications')
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-popup-blocking")
    # New Here
    options.add_argument('--disable-gpu')
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-setuid-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # options.add_argument("user-data-dir={}".format(profile))
    # options.add_argument("--incognito")
    # =============================================================================
    # End items
    # =============================================================================
    driver = uc.Chrome(options=options, use_subprocess=True,version_main = 116)
    
    browser_status = True
    
except Exception as e:
    browser_status = False
    print("Browser Open Fail! : ", e)
    

if browser_status == True:
    
    # ------------------------------
    # Declaring the Empty Arrays
    form_file_txt = []
    field_txt = []
    reporting_txt = []
    filling_entity_txt = []
    cik_txt = []
    located_txt = []
    incorporated_txt = []
    file_no_txt = []
    film_no_txt = []
    
    for i in range(1,11):
        sleep(5)
        
        driver.get(f'https://www.sec.gov/edgar/search/#/q=PIPE&category=custom&forms=S-1&page={i}')
        
        # driver.get('https://www.sec.gov/edgar/search/#/q=PIPE&category=custom&forms=S-1&page=1')
        
        sleep(5)
        # ==============================
        # Click on Check Boxes Done
        driver.find_element(By.ID, value="col-cik").click()
        driver.find_element(By.ID, value="col-located").click()
        driver.find_element(By.ID, value="col-incorporated").click()
        driver.find_element(By.ID, value="col-file-num").click()
        driver.find_element(By.ID, value="col-film-num").click()
        
        sleep(3)
        
        
        
        
        
        # ---------------------------------
        # Form File Data Scrape
        form_file = driver.find_elements(By.CLASS_NAME, value="filetype")
        
        for i in form_file:
            form_file_txt.append(i.text)
            print(i.text)
            
            
        # ----------------------------------
        # Field Data Scrape
        field = driver.find_elements(By.CLASS_NAME, value="filed")
        
        for i in field:
            field_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # reporting Data Scrape
        reporting = driver.find_elements(By.CLASS_NAME, value="enddate")
        
        for i in reporting:
            reporting_txt.append(i.text)
            print(i.text)
        
        # ----------------------------------
        # filling_entity Data Scrape
        filling_entity = driver.find_elements(By.CLASS_NAME, value="entity-name")
        
        for i in filling_entity:
            filling_entity_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # cik Data Scrape
        cik = driver.find_elements(By.CLASS_NAME, value="cik")
        
        for i in cik:
            cik_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # located Data Scrape
        located = driver.find_elements(By.CLASS_NAME, value="biz-location")
        
        for i in located:
            located_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # incorporated Data Scrape
        incorporated = driver.find_elements(By.CLASS_NAME, value="incorporated")
        
        for i in incorporated:
            incorporated_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # file_no Data Scrape
        file_no = driver.find_elements(By.CLASS_NAME, value="file-num")
        
        for i in file_no:
            file_no_txt.append(i.text)
            print(i.text)
            
        # ----------------------------------
        # film_no Data Scrape
        film_no = driver.find_elements(By.CLASS_NAME, value="film-num")
        
        for i in film_no:
            film_no_txt.append(i.text)
            print(i.text)
        
        del form_file_txt[0]
        del field_txt[0]
        del reporting_txt[0]
        del filling_entity_txt[0]
        del cik_txt[0]
        del incorporated_txt[0]
        del file_no_txt[0]
        del film_no_txt[0]
        
        
print(len(form_file_txt))
print(len(field_txt))
print(len(reporting_txt))
print(len(filling_entity_txt))
print(len(cik_txt))
print(len(located_txt))
print(len(incorporated_txt))
print(len(file_no_txt))
print(len(film_no_txt))

dictionary = {
"Form & File" : form_file_txt,
"Filed" : field_txt,
"Reporting for" : reporting_txt,
"Filling entity/person" : filling_entity_txt,
"CIK" : cik_txt,
"Located" : located_txt,
"Incorporated" : incorporated_txt,
"File number" : file_no_txt,
"Film number" : film_no_txt
}
df = pd.DataFrame(dictionary)

excel_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "excel_path")

df.to_excel(r'C:\Users\SA\Desktop\Scrapping\File_Name.xlsx', index = False)
        
    # //*[@id="hits"]/table/tbody/tr[1]/td[5]
    
    # field1 = driver.find_element(By.CLASS_NAME, value="page-link")
    # field1.click()
    # for i in field1:
    #     print(i.text)