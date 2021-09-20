import requests
import os
from bs4 import BeautifulSoup


##Section 1: Define Pre Paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research Report Automation Project') 
data_location                  = os.path.join(project_location,'Data')
costar_summary_location        = os.path.join(data_location,'Costar Summaries') 


summary_file = open(os.path.join(costar_summary_location,'LA.html'),"r")
html_file = summary_file.read()
                


soup = BeautifulSoup(html_file, 'html.parser')

sections = soup.findAll("div", class_="cscc-detail-narrative__title")

for section in sections:
    print(section)
    matches = soup.findAll("div", class_="cscc-narrative-text")
    for m in matches:
        print(m)
        print('-'*20)
        


    
summary_file.close()
