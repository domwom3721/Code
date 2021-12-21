import csv
import os
from numpy import float64
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
#print(df_salesforce)
df = pd.DataFrame(df_salesforce)
#print (df.dtypes)

df['lat_long'] = df["Property Latitude"].astype(str) + ',' + df['Property Longitude'].astype(str)
df.head()
print(df_salesforce)
