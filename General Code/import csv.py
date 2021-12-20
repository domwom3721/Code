import csv
import os
import pandas as pd
from pandas import read_csv

#Define file location pre paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')  
census_data_location           =  os.path.join(project_location,'Data','Neighborhood Reports Data','Census Area Codes') 

#Define location of raw Census Places data files
raw_Census_places_file           =  os.path.join(census_data_location,'national_places.csv')

# Import raw census places data as pandas data frames
open(raw_Census_places_file,'rb')
#df_places = pd.read_csv(raw_Census_places_file)
#print(df_places)