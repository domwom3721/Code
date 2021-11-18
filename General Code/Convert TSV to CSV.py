import pandas as pd
import os

tsv_file = os.path.join(os.environ['USERPROFILE'], 'Desktop','weekly_housing_market_data_most_recent.tsv') 
csv_file = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project','Data','Realtor Writeups','weekly_housing_market_data_most_recent.csv') 
csv_table=pd.read_table(tsv_file,sep='\t')
csv_table.to_csv(csv_file,index=False)
