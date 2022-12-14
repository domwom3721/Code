#Date: 5/2/2022
#Author: Research 2022 Q2
#Summary: Combines multiple RedFin residential real estate data files togeteher

import os
from unicodedata import numeric
from numpy import datetime64
import pandas as pd
#Define file pre-paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research','Projects','Research Report Automation Project') 
raw_data_location              =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Raw')
clean_data_location            =  os.path.join(project_location,'Data\Residential Reports Data\RedFin Data\Clean') 


#Start with a blank dataframe we will append
df_master = pd.DataFrame({'Type':[],
                          'Region':[],})


#From RedFin, for the following geographic levels (State, Metro, County, and Cities (places), we download a main data file and a corresponding price per sqft file
#We do this for condos only and Single-family total.
for condo_or_sf in ['condo', 'sfr', 'Townhouse']:
    for geographic_level in ['national', 'state', 'metro', 'county', 'place']:

        data_file_path = os.path.join(raw_data_location, (geographic_level + '_' + condo_or_sf + '.csv'))
        ppsf_file_path = os.path.join(raw_data_location, (geographic_level + '_' + condo_or_sf + '_ppsf.csv'))
        
        if (os.path.exists(data_file_path) == False) or (os.path.exists(ppsf_file_path) == False ):
            print('Skipping ',geographic_level, ' ', condo_or_sf )
            continue

        if condo_or_sf == 'condo':
            geo_type = 'Condo'
        elif condo_or_sf == 'sfr':
            geo_type = 'Single Family'
        #elif condo_or_sf == 'Townhouse':    
            #geo_type = 'Townhouse'

        #Read in the main data file
        df = pd.read_csv(data_file_path, 
                           encoding='UTF-8', 
                          sep="\t",
                          #engine = 'python',
                        
                           dtype={'Type': str,
                                'Region':str,
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
                                parse_dates=['Month of Period End'],
                            )
        df['Type'] = geo_type

        #Now Read in the price per square foot file
        df_ppsf = pd.read_csv(ppsf_file_path, 
                             encoding='UTF-8', 
                             sep="\t",
                             header=1,    
                             thousands=','         
                            )
        
        #The ppsf file has a column for each month, we need to convert this data so that each month has a row
        df_ppsf['Type'] = geo_type
        df_ppsf = pd.melt(df_ppsf,  ['Type', 'Region'], value_name='Median Price/SF', var_name='Month of Period End')
        
        #Format our Month variable
        df_ppsf['Month of Period End'] = df_ppsf['Month of Period End'].astype(datetime64) 
        
        #Calculate the YoY % growth in median sale price/SF
        df_ppsf['Year Ago Median Sale Price/SF']  = df_ppsf.groupby(['Type', 'Region'])['Median Price/SF'].shift(12) 
        df_ppsf['Month Ago Median Sale Price/SF'] = df_ppsf.groupby(['Type', 'Region'])['Median Price/SF'].shift(1) 

        df_ppsf['Median Price/SF YoY '] = (((df_ppsf['Median Price/SF']/df_ppsf['Year Ago Median Sale Price/SF']) - 1) * 100)
        df_ppsf['Median Price/SF MoM '] = (((df_ppsf['Median Price/SF']/df_ppsf['Month Ago Median Sale Price/SF']) - 1) * 100)

        #Drop the lagged median price/sf variable    
        df_ppsf = df_ppsf.drop(columns=['Year Ago Median Sale Price/SF', 'Month Ago Median Sale Price/SF'])

        #Now we can merge the main data and the price per sqft data
        df = pd.merge(df, df_ppsf, on=(['Type','Month of Period End','Region']), how='left')
        df['Region Type'] = geographic_level.title()
        df_master = df_master.append(df)




#Clean master df
for col_name in df_master.columns[3:]:
    if col_name == ('Region Type') or col_name == ('YoY Median Sale Price/SF Growth'):
        continue

    #Dont need to clean the columns that are already numeric
    if (df_master[col_name].dtype != str) and  (df_master[col_name].dtype != object):
        continue
    
    df_master[col_name] = df_master[col_name].str.replace('$','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace('%','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace(',','',regex=False)
    df_master[col_name] = df_master[col_name].str.replace('K','',regex=False)


    df_master[col_name] = df_master[col_name].astype(float)
    
    if col_name == 'Median Sale Price':
        df_master[col_name] = df_master[col_name] * 1000

df_master['Unique Subject Name'] = df_master['Type'] + ' - ' +  df_master['Region Type'] + ' - ' +  df_master['Region']

#The variables have inconsistent formatting so we fix a few manually
df_master                        = df_master.rename(columns={' Inventory YoY ': "Inventory YoY "})
df_master                        = df_master.rename(columns={'Days on Market YoY': "Days on Market YoY "})
df_master                        = df_master.rename(columns={'Days on Market MoM': "Days on Market MoM "})


#Export the master df as csv file
df_master.to_csv(os.path.join(clean_data_location,'Clean RedFin Data.csv'),index=False)