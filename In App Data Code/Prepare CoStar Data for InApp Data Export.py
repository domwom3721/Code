#Summary: This program reads 4 csv files from CoStar.com, 1 for each major proprety type. 
#         Each file contains quarterly summary statistics on different markets and submarkets.
#         It appends them together and does some light manipulation as well as creating some new variables and exports a clean file.
#         This clean file will then be pasted directly into a Google Sheet that is connected to the Bowery Appraisal WebApp. 
#         This data can then be inserted into appraisal reports on the WebApp.
#Author: Mike Leahy 04/15/2022

#Import packages we will be using

import os
import pandas as pd
import re

#Define file location pre paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)')  
costar_data_location           =  os.path.join(dropbox_root, 'Research','Projects','App','Data')  

#Import raw CoStar data as pandas dataframes
df_combined     = pd.concat([ pd.read_excel(os.path.join(costar_data_location,'multifamily_raw.xlsx')),
                             pd.read_excel(os.path.join(costar_data_location,'office_raw.xlsx')), 
                             pd.read_excel(os.path.join(costar_data_location,'retail_raw.xlsx')),
                            pd.read_excel(os.path.join(costar_data_location,'industrial_raw.xlsx'))
                             ],ignore_index=True)

#Define data cleaning functions
def DropClustersAndLocation(df): 
    #Drops rows that report data on the cluster geography type
    df = df.loc[df['Geography Type'] != 'Cluster']
    df = df.loc[df['Geography Type'] != 'Location Type:Urban']
    df = df.loc[df['Geography Type'] != 'Location Type:CBD']
    df = df.loc[df['Geography Type'] != 'Location Type:Suburban']
    return(df.copy())

def DropColumns(df): 
    columns_to_drop = ['Slice', 'As Of']
    df              = df.drop(columns=columns_to_drop)
    return(df.copy())

def MetroToMarketAndMarketResearchName(df):
    df['Geography Type']            = df['Geography Type'].str.replace('Metro', 'Market', regex=False)
    df['Property Class Name']       = df['Property Class Name'].str.replace('Multi-Family', 'Multifamily', regex=False)
    df['Market Research Name']      = '' 
    

    #Create Market research name for markets. Example: ("Albany - NY" --- > "NY - Albany - Office")
    df['Market Research Name'].loc[df['Geography Type']=='Market']          = df['Geography Name'].str[-2:] + ' - ' + df['Geography Name'].str[:-5] + ' - ' + df['Property Class Name']  
    
    #Create Market research name for nation. Example: ("United States of America" --- > "US - United States of America - Multifamily" )
    df['Market Research Name'].loc[df['Geography Type']=='National']          = 'US' + ' - ' + 'United States of America' + ' - ' + df['Property Class Name']  
    
    #Create Market research name for submarkets.  Example: ("Boston - MA - Quincy/Braintree" --- > "MA - Quincy/Braintree - Multifamily" )
    df['Market Research Name'].loc[df['Geography Type']=='Submarket']       =  df['Geography Name'].str.extract(r'( [A-Z][A-Z] )',expand = False).str.strip() + ' - '  + df['Geography Name'].str.replace(r'(\w+ - [A-Z][A-Z] - )','',regex = True).str.strip() + ' - ' + df['Property Class Name'] 

    #Create a new column that combines the market research name and CoStar Metric 
    df['Market Research Name & Metric']       = df['Market Research Name'] + ' - ' + df['Concept Name'] 
    
    #Move it to first column
    df                                        = df[ ['Market Research Name & Metric'] + [ col for col in df.columns if col != 'Market Research Name & Metric' ]]
    
    #Replace the original Geography name variable with our constructed market research name
    df['Geography Name']                      = df['Market Research Name']
    df                                        = df.drop(columns=['Market Research Name'])
    
    return(df.copy())

def RenameVariable(df):
    df = df.rename(columns={"Property Class Name": "Property Type", })
    return(df.copy())

#Drop columns
df_combined = DropColumns(df=df_combined)
df_combined = DropClustersAndLocation(df=df_combined)
df_combined = MetroToMarketAndMarketResearchName(df=df_combined)
df_combined = RenameVariable(df=df_combined)

#Append our cleaned dataframes together
df_combined.to_excel(os.path.join(costar_data_location,'InAppData.xlsx'),index=False)
