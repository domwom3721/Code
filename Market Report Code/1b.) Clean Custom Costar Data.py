#Cleans the raw data we download from CoStar for custom jobs (eg: a county retail report for an area outside any markets)
#By Mike Leahy 04/22/2021

#Import packages we will be using
import os
from tkinter.constants import E
import pandas as pd
import numpy as np
import shutil


#Define file location pre paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')  
costar_data_location           =  os.path.join(project_location,'Data','CoStar Data') 

#Define location of raw CoStar data files
raw_custom_file_downloads      =  os.path.join(os.environ['USERPROFILE'], 'Downloads','CommercialDataGrid.xlsx') 
raw_custom_file_downloads_mf   =  os.path.join(os.environ['USERPROFILE'], 'Downloads','MultifamilyDataGrid.xlsx') 
raw_custom_file                =  os.path.join(costar_data_location,'Raw Data','CommercialDataGrid.xlsx')
raw_custom_file_mf             =  os.path.join(costar_data_location,'Raw Data','MultifamilyDataGrid.xlsx') 

clean_custom_file              =  os.path.join(costar_data_location,'Clean Data','Clean Custom CoStar Data.xlsx') 

#move exported data from downloads fodler into data folder
if os.path.exists(raw_custom_file_downloads):
    shutil.move(raw_custom_file_downloads, raw_custom_file)
elif os.path.exists(raw_custom_file_downloads_mf):
    shutil.move(raw_custom_file_downloads_mf, raw_custom_file_mf)

sector                      = input('Enter the name of the prop type: Multifamily, Office, Industrial, or Retail')
# sector                         = 'Multifamily'

if os.path.exists(raw_custom_file_mf) and sector == 'Multifamily':
    df_custom  = pd.read_excel(raw_custom_file_mf,
                    dtype={'Sales Volume Transactions': object
                        }      ) 

elif os.path.exists(raw_custom_file) and sector != 'Multifamily':
    #Import raw CoStar data as pandas data frames
    df_custom  = pd.read_excel(raw_custom_file,
                    dtype={'Sales Volume Transactions': object
                        }      ) 
    

df_custom['Geography Type'] = 'Metro'
df_custom['Geography Name'] = input('Enter the name of the market with the following format: Abilene - TX')
# df_custom['Geography Name'] = 'Adams County - IL'

# print(df_custom)
	

def StripVarName(df):
    df = df.rename(columns=lambda x: x.strip())
    return(df)

def KeepLast10Years(df,groupbylist): #Cut down to last 10 years
    df = df.groupby(groupbylist).tail(41)
    return(df)

def SortData(df): #Sorts by geography and quarter
    df = df.sort_values(by=['Geography Name','Period'])
    return(df)

def CreateYearAndQuarterVariables(df): #seperates the period variable into 2 components (year and quarter)
    df.loc[:,'Year']           =   df.loc[:,'Period'].str[:4]
    df.loc[:,'Quarter']        =   df.loc[:,'Period'].str[5:]
    return(df)

def DropExtraVariables(df,sector): #Drops the variables we don't use in our analysis
    if sector == "Multifamily":
        df = df.drop(columns=['Cap Rate Transactions',
                              'Market Asking Rent Index',
                              'Forecast Scenario'
                              ] 
                    )
    else:
        df = df.drop(columns=['Cap Rate Transactions',
                              'Forecast Scenario'
                             ]
                    )
    return(df)

def CleanNetAbsorption(df,sector):
    if sector != 'Multifamily':
        df.loc[df['Net Absorption SF'] == '-', 'Net Absorption SF'] = 0
        df['Net Absorption SF']        = pd.to_numeric(df['Net Absorption SF'])
        return(df)
    
def DestringVariablesConvertToNumeric(df,sector):

    if sector == 'Multifamily':
        vars_list_to_destring = ['Average Sale Price',
                'Market Cap Rate',
                'Vacancy Rate', 
                'Asset Value',
                'Total Sales Volume',
                'Cap Rate',
                'Existing Buildings',
                'Market Sale Price Growth',
                'Occupancy Rate',
                'Median Cap Rate',
                'Under Construction Buildings',
                'Year',
                'Market Effective Rent/Unit',
                'Market Effective Rent Growth 12 Mo',
                'Market Effective Rent Growth',
                'Under Construction Units',
                'Inventory Units',
                'Absorption Units',
                'Absorption %',
                'Sales Volume Transactions'
                ]
    else:
        vars_list_to_destring = ['Average Sale Price',
                'Market Cap Rate',
                'Vacancy Rate',
                'Availability Rate', 
                'Asset Value',
                'Total Sales Volume',
                'Cap Rate',
                'Existing Buildings',
                'Market Sale Price Growth',
                'Occupancy Rate',
                'Median Cap Rate',
                'Under Construction Buildings',
                'Year',
                'Sales Volume Transactions',
                'Inventory SF',
                'Under Construction SF',
                'Market Rent/SF',
                'Market Rent Growth',
                'Market Rent Growth 12 Mo',
                'Available SF'
                ]


    for var in vars_list_to_destring:
        try:
           
            if df[var].dtype == 'object': #only do the following for string variables
                df[var] = df[var].str.replace('$', '', regex=False)
                df[var] = df[var].str.replace(',', '', regex=False)
                df[var] = df[var].str.replace('%', '', regex=False)
                df[var] = df[var].str.replace('-', '', regex=False)
                df[var] = pd.to_numeric(df[var])
        except Exception as e:
            pass
    return(df)

def FillBlanksWithZero(df,sector):
    if sector == 'Multifamily':
        var_list_to_replace_blanks = ['Sales Volume Transactions','Total Sales Volume','Under Construction Units',]
    else:
        var_list_to_replace_blanks = ['Sales Volume Transactions','Total Sales Volume','Under Construction SF','Availability Rate','Net Absorption SF']

    for var in var_list_to_replace_blanks:
        df[var] = df[var].fillna(0)

    return(df)

def ConvertPercenttoPercentagePoints(df,sector):
    if sector == 'Multifamily':
        # print(df['Absorption Percent'])
        rate_vars = ['Absorption Percent',
                    'Vacancy Rate',
                    ]
    else:
        rate_vars = ['Vacancy Rate',
                    'Market Cap Rate', ]
    
    for var in rate_vars:
        # print(var)
        df[var] = round((df[var] * 100),2)
    return(df)

def MainClean(df,sector): #Calls all cleaning functions and returns cleaned dataframes
    df = StripVarName(df)
    df = SortData(df)
    df = KeepLast10Years(df,['Geography Name'])
    df = CreateYearAndQuarterVariables(df)
    df['Sales Volume Transactions']  = 0
    if sector != 'Multifamily':
        df  = df.rename(columns={'Vacant Percent % Total': "Vacancy Rate",'Total Available Percent % Total':'Availability Rate','Sales Volume':'Total Sales Volume',
                                'Net Absorption SF Total':'Net Absorption SF',
                                'Rent/SF':'Market Rent/SF',
                                'Price/SF': 'Asset Value',
                               })
       
        
        df['Asset Value'] = df['Asset Value'] * df['Inventory SF']
    else:
        df  = df.rename(columns={'Vacancy Percent': "Vacancy Rate",'Total Available Percent % Total':'Availability Rate','Sales Volume':'Total Sales Volume',
                                'Rent/Unit':'Market Rent/SF',
                                'Sale Price Per Unit': 'Asset Value',
                                 'Effective Rent Per Unit':'Market Effective Rent/Unit'})
        
        df['Asset Value'] = df['Asset Value'] * df['Inventory Units']

    df = DestringVariablesConvertToNumeric(df,sector)
    # df = CleanNetAbsorption(df,sector)
    df = ConvertPercenttoPercentagePoints(df,sector)
    df = FillBlanksWithZero(df,sector)
    return(df)



#Data cleaning

df_custom =  MainClean(df_custom,sector)
# print(df_custom)









#Loop through the 4 dataframes: create variables we will use in our report/figures 
for df in [df_custom]:

    df['Geography Name'] = df['Geography Name'].str.replace('New York City', 'Manhattan', regex=False)

    #Clean the Sqft and Unit variables seperately
    if sector == 'Multifamily':

        #Create laggd variables
        df['Lagged Inventory Units']       = df.groupby('Geography Name')['Inventory Units'].shift(1)
      

        #Create variable for apt absorption rate
        df['Absorption Rate'] = round(  ((df['Absorption Units']/df['Inventory Units']) * 100)  ,2) 

        
        #Create variable for inventory growth rate
        df['Inventory Growth'] = round(((df['Inventory Units'] / df['Lagged Inventory Units']) - 1)  * 100,2)

        #Create variable for percent under construction
        df['Under Construction %'] = (df['Under Construction Units']/df['Inventory Units'] ) *100

        #Asset Value per unit
        df['Asset Value/Unit']     = round((df['Asset Value']/df['Inventory Units']),2)

        df['Previous Quarter Asset Value/Unit'] = df.groupby('Geography Name')['Asset Value/Unit'].shift(1)
        df['4 Quarters Ago Asset Value/Unit']   = df.groupby('Geography Name')['Asset Value/Unit'].shift(4)

        df['QoQ Asset Value/Unit Growth']        = round( (((df['Asset Value/Unit']  / df['Previous Quarter Asset Value/Unit']) - 1) * 100),                    1)
        df['YoY Asset Value/Unit Growth']        = round( (((df['Asset Value/Unit']  / df['4 Quarters Ago Asset Value/Unit'])   - 1) * 100),                    1)

        #Market Rent
        df['Previous Quarter Market Effective Rent/Unit'] = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(1)
        df['4 Quarters Ago Market Effective Rent/Unit']   = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(4)

        df['QoQ Market Effective Rent/Unit Growth']        = round( (((df['Market Effective Rent/Unit']   / df['Previous Quarter Market Effective Rent/Unit']) - 1) * 100),                    1)
        df['YoY Market Effective Rent/Unit Growth']        = round( (((df['Market Effective Rent/Unit']  / df['4 Quarters Ago Market Effective Rent/Unit'])   - 1) * 100),                    1)
        
        #Absorption Units
        df['Previous Quarter Absorption Units'] = df.groupby('Geography Name')['Absorption Units'].shift(1)
        df['4 Quarters Ago Absorption Units']   = df.groupby('Geography Name')['Absorption Units'].shift(4)

        # df['QoQ Absorption Units Growth']        = round( (((df['Absorption Units']   / abs(df['Previous Quarter Absorption Units'])) - 1) * 100),                    1)
        # df['YoY Absorption Units Growth']        = round( (((df['Absorption Units']  / abs(df['4 Quarters Ago Absorption Units']))   - 1) * 100),                    1)
           
        df['QoQ Absorption Units Growth']        = round((df['Absorption Units']   - df['Previous Quarter Absorption Units'])    / abs(df['Previous Quarter Absorption Units'])  * 100,1)              
        df['YoY Absorption Units Growth']        = round((df['Absorption Units']   - df['4 Quarters Ago Absorption Units'])      /  abs(df['4 Quarters Ago Absorption Units'])   * 100 ,1)           
           

    else:            
        #Create laggd variables
        df['Lagged Inventory SF']       = df.groupby('Geography Name')['Inventory SF'].shift(1)


        #Create variable for absorption rate 
        df['Net Absorption SF']         = pd.to_numeric(df['Net Absorption SF'])
        df['Absorption Rate']           = round((df['Net Absorption SF'] / df['Inventory SF']) * 100,2 )

                
        #Absorption SF
        df['Previous Quarter Net Absorption SF'] = df.groupby('Geography Name')['Net Absorption SF'].shift(1)
        df['4 Quarters Ago Net Absorption SF']   = df.groupby('Geography Name')['Net Absorption SF'].shift(4)

        # df['QoQ Net Absorption SF Growth']        = round( (((df['Net Absorption SF']   / abs(df['Previous Quarter Net Absorption SF'])) - 1) * 100),                    1)
        # df['YoY Net Absorption SF Growth']        = round( (((df['Net Absorption SF']  / abs(df['4 Quarters Ago Net Absorption SF']))   - 1) * 100),                    1)
        
        df['QoQ Net Absorption SF Growth']        = round((df['Net Absorption SF']   - df['Previous Quarter Net Absorption SF'])    / abs(df['Previous Quarter Net Absorption SF'])  * 100,1)              
        df['YoY Net Absorption SF Growth']        = round((df['Net Absorption SF']   - df['4 Quarters Ago Net Absorption SF'])      /  abs(df['4 Quarters Ago Net Absorption SF'])   * 100 ,1)           
           

        #Availability Rate 
        df['Previous Quarter Availability Rate'] = df.groupby('Geography Name')['Availability Rate'].shift(1)
        df['4 Quarters Ago Availability Rate']   = df.groupby('Geography Name')['Availability Rate'].shift(4)
    
        df['QoQ Availability Rate Growth']        = round((df['Availability Rate'] - df['Previous Quarter Availability Rate']) * 100,0)
        df['YoY Availability Rate Growth']        = round((df['Availability Rate'] - df['4 Quarters Ago Availability Rate'])   * 100,0)

        df['Availability Rate']                   = round(df['Availability Rate'],1)



        #Market Rent
        df['Previous Quarter Market Rent/SF'] = df.groupby('Geography Name')['Market Rent/SF'].shift(1)
        df['4 Quarters Ago Market Rent/SF']   = df.groupby('Geography Name')['Market Rent/SF'].shift(4)

        df['QoQ Rent Growth']        = round( (((df['Market Rent/SF']  / df['Previous Quarter Market Rent/SF']) - 1) * 100),                    1)
        df['YoY Rent Growth']        = round( (((df['Market Rent/SF']  / df['4 Quarters Ago Market Rent/SF'])   - 1) * 100),                    1)

        
        #Create variable for inventory growth rate
        df['Inventory Growth'] = round(((df['Inventory SF'] / df['Lagged Inventory SF']) - 1)  * 100,2)


        #Create variable for percent under construction
        df['Under Construction %'] = (df['Under Construction SF']/df['Inventory SF'] ) *100

        #Asset Value per sqft
        df['Asset Value/Sqft']     = round((df['Asset Value']/df['Inventory SF']),2)

        df['Previous Quarter Asset Value/Sqft'] = df.groupby('Geography Name')['Asset Value/Sqft'].shift(1)
        df['4 Quarters Ago Asset Value/Sqft']   = df.groupby('Geography Name')['Asset Value/Sqft'].shift(4)

        df['QoQ Asset Value/Sqft Growth']        = round( (((df['Asset Value/Sqft']  / df['Previous Quarter Asset Value/Sqft']) - 1) * 100),                    1)
        df['YoY Asset Value/Sqft Growth']        = round( (((df['Asset Value/Sqft']  / df['4 Quarters Ago Asset Value/Sqft'])   - 1) * 100),                    1)



    #Making Variables for all sectors
    df['Previous Quarter Vacancy'] = df.groupby('Geography Name')['Vacancy Rate'].shift(1)
    df['4 Quarters Ago Vacancy']   = df.groupby('Geography Name')['Vacancy Rate'].shift(4)
    
    df['QoQ Vacancy Growth']        = round((df['Vacancy Rate'] - df['Previous Quarter Vacancy']) * 100,0)
    df['YoY Vacancy Growth']        = round((df['Vacancy Rate'] - df['4 Quarters Ago Vacancy'])   * 100,0)

    #Absorption
    df['Previous Quarter Absorption Rate'] =  df.groupby('Geography Name')['Absorption Rate'].shift(1)
    df['4 Quarters Ago Absorption Rate']   =  df.groupby('Geography Name')['Absorption Rate'].shift(4)

    df['QoQ Absorption Growth']        = round((df['Absorption Rate'] - df['Previous Quarter Absorption Rate']) * 100,0)
    df['YoY Absorption Growth']        = round((df['Absorption Rate'] - df['4 Quarters Ago Absorption Rate'])   * 100,0)
    
    #Sales Volume
    df['Previous Quarter Total Sales Volume'] = df.groupby('Geography Name')['Total Sales Volume'].shift(1)
    df['4 Quarters Ago Total Sales Volume']   = df.groupby('Geography Name')['Total Sales Volume'].shift(4)
    
    df['QoQ Total Sales Volume Growth']        = round( (((df['Total Sales Volume']  / df['Previous Quarter Total Sales Volume']) - 1) * 100),                    0)
    df['YoY Total Sales Volume Growth']        = round( (((df['Total Sales Volume']  / df['4 Quarters Ago Total Sales Volume'])   - 1) * 100),                    0)

    #Transactions
    df['Previous Quarter Sales Volume Transactions'] = df.groupby('Geography Name')['Sales Volume Transactions'].shift(1)
    df['4 Quarters Ago Sales Volume Transactions']   = df.groupby('Geography Name')['Sales Volume Transactions'].shift(4)
    
    df['QoQ Transactions Growth']         = round(  (((df['Sales Volume Transactions']/df['Previous Quarter Sales Volume Transactions']) - 1)  * 100)            ,0)
    df['YoY Transactions Growth']         = round(  (((df['Sales Volume Transactions']/df['4 Quarters Ago Sales Volume Transactions'])   - 1)  * 100)            ,0)


    #market cap rate
    df['Market Cap Rate']                 = 0
    df['Previous Quarter Market Cap Rate'] = df.groupby('Geography Name')['Market Cap Rate'].shift(1)
    df['4 Quarters Ago Market Cap Rate']   = df.groupby('Geography Name')['Market Cap Rate'].shift(4)
    
    df['QoQ Market Cap Rate Growth']        = round((df['Market Cap Rate'] - df['Previous Quarter Market Cap Rate']) * 100,0)
    df['YoY Market Cap Rate Growth']        = round((df['Market Cap Rate'] - df['4 Quarters Ago Market Cap Rate'])   * 100,0)

    # #Round  3 percentage variables we report in overview table
    # df['Market Cap Rate']            = round(df['Market Cap Rate'],1)
    df['Vacancy Rate']               = round(df['Vacancy Rate'],1)
    df['Absorption Rate']            = round(df['Absorption Rate'],1)



    df_custom.to_excel(clean_custom_file)






