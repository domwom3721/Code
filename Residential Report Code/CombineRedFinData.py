#Date: 5/2/2022
#Author: Mike Leahy
#Summary: Combines multiple RedFin residential real estate data files togeteher

import os
from turtle import st
import pandas as pd

#Define file pre-paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project') 
raw_data_location              =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Raw')
clean_data_location            =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Clean') 



#Start with a blank dataframe we will append
df_master = pd.DataFrame({'Type':[],
                          'Region':[],})


#From RedFin, for the following geographic levels (State, Metro, County, and Cities (places), we download a main data file and a corresponding price per sqft file
#We do this for condos only and Single-family total.
for condo_or_sf in ['condo', 'sf']:
    for geographic_level in ['state', 'metro', 'county', 'place']:

        data_file_path = os.path.join(raw_data_location, (condo_or_sf + '_' +geographic_level+ '_data.csv'))
        ppsf_file_path = os.path.join(raw_data_location, (condo_or_sf + '_' +geographic_level+ '_ppsf.csv'))
        
        #Read in the main data file
        df = pd.read_csv(data_file_path, 
                           encoding='UTF-16 LE', 
                          sep="\t",
                        
                           dtype={'Type': str,
                                'Region':str,
                                'Month of Period End':str,
                                'Median Sale Price':str	,
                                'Median Sale Price MoM':str ,	
                                'Median Sale Price YoY':str ,	
                                'Homes Sold':str	,
                                'Homes Sold MoM':str ,
                                'Homes Sold YoY':str ,	
                                'New Listings':str	,
                                'New Listings MoM':str , 	
                                'New Listings YoY':str ,	
                                'Inventory':str,
                                'Inventory MoM':str ,	 
                                'Inventory YoY':str ,	
                                'Days on Market':str,	
                                'Days on Market MoM':str ,	
                                'Days on Market YoY':str	,
                                'Average Sale To List':str ,
                                'Average Sale To List MoM':str ,	
                                'Average Sale To List YoY':str ,
                                },                         
                            )
        
        #Now Read in the price per square foot file
        df_ppsf = pd.read_csv(ppsf_file_path, 
                             encoding='UTF-16 LE', 
                             sep="\t",
                             header=1         
                            )
        

        #The ppsf file has a column for each month, we need to convert this data so that each month has a row
        df_ppsf = pd.melt(df_ppsf,  ['Type', 'Region'], value_name='Price Per Sqft', var_name='Month of Period End')
        

        #Now we can merge the main data and the price per sqft data
        df = pd.merge(df, df_ppsf, on=(['Type','Month of Period End','Region']), how='left')
        
        df_master = df_master.append(df)

#Clean master df
for col_name in df_master.columns[3:]:
    
    df_master[col_name] = df_master[col_name].str.replace('$','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace('%','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace(',','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace('K','',regex=False)


    df_master[col_name] = df_master[col_name].astype(float)
    
    if col_name == 'Median Sale Price':
        df_master[col_name] = df_master[col_name] * 100000

#Export the master df as csv file
df_master.to_csv(os.path.join(clean_data_location,'Clean RedFin Data.csv'),index=False)