import os
import pandas as pd
#Author: Mike Leahy
#Date: 01/19/2022
#Summary: This script injests data on key metrics for different hotel markets across the country from CoStar.com.
          #It cleans the data and exports a clean file

#Define file pre paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project')      #Main Folder that stores all output, code, and documentation
hotel_data_location            = os.path.join(project_location,'Data','Hotel Reports Data')                                   #Folder with data for hotel reports only


#Step 1: Import Raw Data
raw_hotel_df                 = pd.read_csv(os.path.join(hotel_data_location,'Raw Hotel Data','hotel.csv')) 

#Step 2: Clean Data
def MainClean(df):
    return(df)

clean_hotel_df  = MainClean(df = raw_hotel_df)

#Step 3: Export Cleaned Data
clean_hotel_df.to_csv(os.path.join(hotel_data_location,'Clean Hotel Data','clean_hotel_data.csv'))