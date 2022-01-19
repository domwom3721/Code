#Author: Mike Leahy
#Date: 01/19/2022
#Summary: This script injests data on key metrics for different hotel markets across the country from CoStar.com.
          #It creates a report document for each market and saves it in a corresponding folder

import os
import pandas as pd

#Define file pre paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project')      #Main Folder that stores all output, code, and documentation
# output_location                = os.path.join(dropbox_root,'Research','Market Analysis','Market','Other','Hotel')           #The folder where we store our current reports, production
output_location                = os.path.join(project_location,'Output','Hotel')                                              #The folder where we store our current reports, testing folder
map_location                   = os.path.join(project_location,'Data','Hotel Reports Data','CoStar Maps')                     #Folders with maps png files  
general_data_location          =  os.path.join(project_location,'Data','General Data')                                        #Folder with data used in multiple report types
hotel_data_location            = os.path.join(project_location,'Data','Hotel Reports Data')                                   #Folder with data for hotel reports only


#Import hotel data
hotel_df                 = pd.read_csv(os.path.join(hotel_data_location,'Clean Hotel Data','clean_hotel_data.csv')) 
 
print(hotel_df)