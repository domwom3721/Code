import csv
import os
import pandas as pd
from pandas import read_csv
from pandas.core.frame import DataFrame

#Define file location pre paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')  
salesforce_report              =  os.path.join(project_location,'Data','Neighborhood Reports Data','Salesforce') 

#Define location of raw Census Places data files
daily_salesforce_jobs          =  os.path.join(salesforce_report,'report.csv')

# Import raw census places data as pandas data frames
open(daily_salesforce_jobs,'rb')
df_salesforce = pd.read_csv(daily_salesforce_jobs)
print(df_salesforce)

#df_salesforce["latlong"] = df_salesforce['Property Latitude', 'Property Longitude'].agg(','.join)
#df_salesforce["lat_long"] = list(zip(df_salesforce["property_latitude"], df_salesforce["property_longitude"]))
#print(df_salesforce)