#Date: 4/18/2022
#Author: Mike Leahy
#Summary: Attempt at scraping the list of MIA designee holders from https://ai.appraisalinstitute.org/eweb/DynamicPage.aspx?webcode=aifaasearch
#         The purpose is to help guide Bowery determine which markets to expand into it

import requests
import selenium
import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup

#Step 1: Open a browser and open the website
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option('useAutomationExtension', False)
browser = webdriver.Chrome(service=Service(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop', 'chromedriver.exe'))), options=options)
browser.get('https://ai.appraisalinstitute.org/eweb/DynamicPage.aspx?webcode=aifaasearch')
time.sleep(10)

# Step 2: Select the MIA designation
designation_box = Select(browser.find_element(By.NAME, "ai_designation"))
designation_box.select_by_visible_text('MAI')

#Step 3: Hit Search Button
search_button = browser.find_element(By.ID, "searchButton")
search_button.click()


def process_rows(rows):
    for count, row in enumerate(rows):
        # print(row)
        info = row.text
        name = info.split(',',2)[0]
        info = info.replace(name + ', ', '')
        info = info[info.find('Member') + len('Member'):]
        city = info.split(',',2)[0]
        info = info.replace(city + ', ', '')
        state = info[0:2]
        info  = info[2:]
        print(info)
        email = info
        address = info
        phone_number = info
        if count > 0:
            names.append(name) 
            emails.append(email) 
            states.append(state) 
            cities.append(city) 
            phone_numbers.append(phone_number) 
            addresses.append(address)   

#Step 4: Get HTML for page 1
time.sleep(20)
html        = browser.page_source
soup_data   = BeautifulSoup(html, 'html.parser')
rows       = soup_data.findAll("tr", {"role":"row"})

#Step 5: Create Empty lists we will fill with info from each row on each page
names  = []
emails = []
states = []
cities = []
phone_numbers = []
addresses     = []

process_rows(rows = rows)
time.sleep(5)



# #Step 5: Keep hitting next page save the html and add the info to our lists
# while True:
#     next_page_button = browser.find_element(By.ID, "faaSearchResults_next")
#     time.sleep(3)
#     next_page_button.click()
#     process_rows(rows = BeautifulSoup(browser.page_source, 'html.parser').findAll("tr", {"role":"row"}))
#     time.sleep(5)   


df = pd.DataFrame({'Name':names,
                'Email':emails,
                'Adress': addresses,
                'State':         states,
                "City":cities,
                'Phone Number':phone_numbers,
                })
df = df.sort_values(by=['State','City'], ascending = (True, True))
print(df)

df.to_excel(os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects', 'Research Report Automation Project','Output','MAI','MAI.xlsx'), index=False )