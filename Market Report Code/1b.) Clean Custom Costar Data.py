#Cleans the raw data we download from CoStar for custom jobs (eg: a county retail report for an area outside any markets)
#By Mike Leahy 04/22/2021

#Import packages we will be using
import os
from tkinter.constants import E
import pandas as pd
import numpy as np
import shutil

#Get user input for sector and geography
sector                         = input('Enter the name of the prop type: Multifamily, Office, Industrial, or Retail (m/o/r/i)')

while (sector != 'm' ) and (sector != 'i' ) and (sector != 'o' )  and (sector != 'r' ):
    print('Not an accepted sector, try again')
    sector                         = input('Enter the name of the prop type: Multifamily, Office, Industrial, or Retail (m/o/r/i')

if sector == 'm':
    sector = 'Multifamily'
elif sector == 'o':
    sector = 'Office'
elif sector == 'i':
    sector = 'Industrial'
elif sector == 'r':
    sector = 'Retail'

#geography_name                 = input('Enter the name of the market with the following format: Abilene - TX')
geography_name                 = 'Example - NY'


#Define file location pre paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')  
costar_data_location           =  os.path.join(project_location,'Data','CoStar Data') 

#Define the location of the downloaded files and where we want to move them to
if sector != 'Multifamily':
    raw_download_data_file            = os.path.join(os.environ['USERPROFILE'], 'Downloads','CommercialDataGrid.xlsx') 
    raw_main_data_file                = os.path.join(costar_data_location,'Raw Data','Custom County Data','CommercialDataGrid.xlsx')
    
    raw_download_sales_volume_file    =  os.path.join(os.environ['USERPROFILE'], 'Downloads','Sales Volume & Market Sale Price Per SF.xlsx') 
    raw_sales_volume_file             =  os.path.join(costar_data_location,'Raw Data','Custom County Data','Sales Volume & Market Sale Price Per SF.xlsx')
    
    raw_download_market_cap_rate_file = os.path.join(os.environ['USERPROFILE'], 'Downloads','Market Cap Rate.xlsx') 
    raw_market_cap_rate_file          = os.path.join(costar_data_location,'Raw Data','Custom County Data','Market Cap Rate.xlsx')

    raw_download_market_rent_file     = os.path.join(os.environ['USERPROFILE'], 'Downloads','Market Rent Per SF.xlsx') 
    raw_market_rent_file              = os.path.join(costar_data_location,'Raw Data','Custom County Data','Market Rent Per SF.xlsx')
    
    clean_main_data_file              = os.path.join(costar_data_location,'Clean Data','retail_clean.csv')

else:
    raw_download_data_file            = os.path.join(os.environ['USERPROFILE'], 'Downloads','MultifamilyDataGrid.xlsx') 
    raw_main_data_file                = os.path.join(costar_data_location,'Raw Data','Custom County Data','MultifamilyDataGrid.xlsx') 
    
    raw_download_sales_volume_file    = os.path.join(os.environ['USERPROFILE'], 'Downloads','Sales Volume & Market Sale Price Per Unit.xlsx') 
    raw_sales_volume_file             = os.path.join(costar_data_location,'Raw Data','Custom County Data','Sales Volume & Market Sale Price Per Unit.xlsx')
    
    raw_download_market_cap_rate_file = os.path.join(os.environ['USERPROFILE'], 'Downloads','Market Cap Rate.xlsx') 
    raw_market_cap_rate_file          = os.path.join(costar_data_location,'Raw Data','Custom County Data','Market Cap Rate.xlsx')

    raw_download_market_rent_file     = os.path.join(os.environ['USERPROFILE'], 'Downloads','Market Rent Per SF.xlsx') 
    raw_market_rent_file              = os.path.join(costar_data_location,'Raw Data','Custom County Data','Market Rent Per SF.xlsx')
    
    clean_main_data_file              = os.path.join(costar_data_location,'Clean Data','mf_clean.csv')

clean_custom_file                      =  os.path.join(costar_data_location,'Clean Data','Clean Custom CoStar Data.xlsx') 



#Move exported data from downloads fodler into data folder
if os.path.exists(raw_download_data_file):
    shutil.move(raw_download_data_file,raw_main_data_file )

if os.path.exists(raw_download_sales_volume_file):
    shutil.move(raw_download_sales_volume_file, raw_sales_volume_file)

if os.path.exists(raw_download_market_cap_rate_file):
    shutil.move(raw_download_market_cap_rate_file,raw_market_cap_rate_file )

if os.path.exists(raw_download_market_rent_file) and (sector != 'Multifamily'):
    shutil.move(raw_download_market_rent_file,raw_market_rent_file )

#Now our downloaded data files are in the raw data folder, we will merge them together into a single clean file we export
df_custom                            = pd.read_excel(raw_main_data_file)
df_custom_sales_volume               = pd.read_excel(raw_sales_volume_file)
df_custom_market_cap_rate            = pd.read_excel(raw_market_cap_rate_file)
df_clean_file_for_last_period        = pd.read_csv(clean_main_data_file)

#Get the latest period
latest_period = df_clean_file_for_last_period['Period'].iloc[-1]

#For non MF, rename the rent variable
if sector != 'Multifamily':
    df_custom_market_rent            = pd.read_excel(raw_market_rent_file)
    df_custom_market_rent            = df_custom_market_rent.rename(columns={"Current Search": "Market Rent/SF"})


#Start by changing market cap rate variable name
df_custom_market_cap_rate =  df_custom_market_cap_rate.rename(columns={"Current Search": "Market Cap Rate"})

#Remove white space from period variable name in main custom dataframe
df_custom =  df_custom.rename(columns={"  Period": "Period"})

#Merge in the cap rate and sales volume dataframe with the regular custom dataframe
df_custom                 =  pd.merge(df_custom, df_custom_sales_volume, on=['Period'],how = 'left') 
df_custom                 =  pd.merge(df_custom, df_custom_market_cap_rate, on=['Period'],how = 'left') 

#merge in market rent/sf if non-multifamily
if sector != 'Multifamily':
    df_custom                 =  pd.merge(df_custom, df_custom_market_rent, on=['Period'],how = 'left') 


#Sort oldest to newest
df_custom                 =  df_custom.sort_values(by=['Period'])

#Restrict to latest quarter we are doing reports on 
while df_custom['Period'].iloc[-1] != latest_period:
    df_custom = df_custom[0:len(df_custom) -1]



df_custom['Geography Type'] = 'Metro'
df_custom['Geography Name'] = geography_name

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
df_custom['Geography Name'] = df_custom['Geography Name'].str.replace('New York City', 'Manhattan', regex=False)

#Clean the Sqft and Unit variables seperately
if sector == 'Multifamily':

    #Create laggd variables
    df_custom['Lagged Inventory Units']       = df_custom.groupby('Geography Name')['Inventory Units'].shift(1)
    

    #Create variable for apt absorption rate
    df_custom['Absorption Rate'] = round(  ((df_custom['Absorption Units']/df_custom['Inventory Units']) * 100)  ,2) 

    
    #Create variable for inventory growth rate
    df_custom['Inventory Growth'] = round(((df_custom['Inventory Units'] / df_custom['Lagged Inventory Units']) - 1)  * 100,2)

    #Create variable for percent under construction
    df_custom['Under Construction %'] = (df_custom['Under Construction Units']/df_custom['Inventory Units'] ) *100

    #Asset Value per unit
    df_custom['Asset Value/Unit']     = round((df_custom['Asset Value']/df_custom['Inventory Units']),2)

    df_custom['Previous Quarter Asset Value/Unit'] = df_custom.groupby('Geography Name')['Asset Value/Unit'].shift(1)
    df_custom['4 Quarters Ago Asset Value/Unit']   = df_custom.groupby('Geography Name')['Asset Value/Unit'].shift(4)

    df_custom['QoQ Asset Value/Unit Growth']        = round( (((df_custom['Asset Value/Unit']  / df_custom['Previous Quarter Asset Value/Unit']) - 1) * 100),                    1)
    df_custom['YoY Asset Value/Unit Growth']        = round( (((df_custom['Asset Value/Unit']  / df_custom['4 Quarters Ago Asset Value/Unit'])   - 1) * 100),                    1)

    #Market Rent
    df_custom['Previous Quarter Market Effective Rent/Unit'] = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(1)
    df_custom['4 Quarters Ago Market Effective Rent/Unit']   = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(4)

    df_custom['QoQ Market Effective Rent/Unit Growth']        = round( (((df_custom['Market Effective Rent/Unit']   / df_custom['Previous Quarter Market Effective Rent/Unit']) - 1) * 100),                    1)
    df_custom['YoY Market Effective Rent/Unit Growth']        = round( (((df_custom['Market Effective Rent/Unit']  / df_custom['4 Quarters Ago Market Effective Rent/Unit'])   - 1) * 100),                    1)
    
    #Absorption Units
    df_custom['Previous Quarter Absorption Units'] = df_custom.groupby('Geography Name')['Absorption Units'].shift(1)
    df_custom['4 Quarters Ago Absorption Units']   = df_custom.groupby('Geography Name')['Absorption Units'].shift(4)

    df_custom['QoQ Absorption Units Growth']        = round((df_custom['Absorption Units']   - df['Previous Quarter Absorption Units'])    / abs(df_custom['Previous Quarter Absorption Units'])  * 100,1)              
    df_custom['YoY Absorption Units Growth']        = round((df_custom['Absorption Units']   - df['4 Quarters Ago Absorption Units'])      /  abs(df_custom['4 Quarters Ago Absorption Units'])   * 100 ,1)           
        

else:            
    #Create laggd variables
    df_custom['Lagged Inventory SF']       = df_custom.groupby('Geography Name')['Inventory SF'].shift(1)

    #Create variable for absorption rate 
    df_custom['Net Absorption SF']         = pd.to_numeric(df_custom['Net Absorption SF'])
    df_custom['Absorption Rate']           = round((df_custom['Net Absorption SF'] / df_custom['Inventory SF']) * 100,2 )

    #Absorption SF
    df_custom['Previous Quarter Net Absorption SF'] = df_custom.groupby('Geography Name')['Net Absorption SF'].shift(1)
    df_custom['4 Quarters Ago Net Absorption SF']   = df_custom.groupby('Geography Name')['Net Absorption SF'].shift(4)

    df_custom['QoQ Net Absorption SF Growth']        = round((df_custom['Net Absorption SF']   - df_custom['Previous Quarter Net Absorption SF'])    / abs(df_custom['Previous Quarter Net Absorption SF'])  * 100,1)              
    df_custom['YoY Net Absorption SF Growth']        = round((df_custom['Net Absorption SF']   - df_custom['4 Quarters Ago Net Absorption SF'])      /  abs(df_custom['4 Quarters Ago Net Absorption SF'])   * 100 ,1)                 

    #Availability Rate 
    df_custom['Previous Quarter Availability Rate'] = df_custom.groupby('Geography Name')['Availability Rate'].shift(1)
    df_custom['4 Quarters Ago Availability Rate']   = df_custom.groupby('Geography Name')['Availability Rate'].shift(4)

    df_custom['QoQ Availability Rate Growth']        = round((df_custom['Availability Rate'] - df_custom['Previous Quarter Availability Rate']) * 100,0)
    df_custom['YoY Availability Rate Growth']        = round((df_custom['Availability Rate'] - df_custom['4 Quarters Ago Availability Rate'])   * 100,0)
    df_custom['Availability Rate']                   = round(df_custom['Availability Rate'],1)

    #Market Rent
    df_custom['Previous Quarter Market Rent/SF'] = df_custom.groupby('Geography Name')['Market Rent/SF'].shift(1)
    df_custom['4 Quarters Ago Market Rent/SF']   = df_custom.groupby('Geography Name')['Market Rent/SF'].shift(4)

    df_custom['QoQ Rent Growth']        = round( (((df_custom['Market Rent/SF']  / df_custom['Previous Quarter Market Rent/SF']) - 1) * 100),                    1)
    df_custom['YoY Rent Growth']        = round( (((df_custom['Market Rent/SF']  / df_custom['4 Quarters Ago Market Rent/SF'])   - 1) * 100),                    1)

    #Create variable for inventory growth rate
    df_custom['Inventory Growth'] = round(((df_custom['Inventory SF'] / df_custom['Lagged Inventory SF']) - 1)  * 100,2)

    #Create variable for percent under construction
    df_custom['Under Construction %'] = (df_custom['Under Construction SF']/df_custom['Inventory SF'] ) *100

    #Asset Value per sqft
    df_custom['Asset Value/Sqft']     = round((df_custom['Asset Value']/df_custom['Inventory SF']),2)

    df_custom['Previous Quarter Asset Value/Sqft'] = df_custom.groupby('Geography Name')['Asset Value/Sqft'].shift(1)
    df_custom['4 Quarters Ago Asset Value/Sqft']   = df_custom.groupby('Geography Name')['Asset Value/Sqft'].shift(4)

    df_custom['QoQ Asset Value/Sqft Growth']        = round( (((df_custom['Asset Value/Sqft']  / df_custom['Previous Quarter Asset Value/Sqft']) - 1) * 100),                    1)
    df_custom['YoY Asset Value/Sqft Growth']        = round( (((df_custom['Asset Value/Sqft']  / df_custom['4 Quarters Ago Asset Value/Sqft'])   - 1) * 100),                    1)



#Making Variables for all sectors
df_custom['Previous Quarter Vacancy'] = df_custom.groupby('Geography Name')['Vacancy Rate'].shift(1)
df_custom['4 Quarters Ago Vacancy']   = df_custom.groupby('Geography Name')['Vacancy Rate'].shift(4)

df_custom['QoQ Vacancy Growth']        = round((df_custom['Vacancy Rate'] - df_custom['Previous Quarter Vacancy']) * 100,0)
df_custom['YoY Vacancy Growth']        = round((df_custom['Vacancy Rate'] - df_custom['4 Quarters Ago Vacancy'])   * 100,0)

#Absorption
df_custom['Previous Quarter Absorption Rate'] =  df_custom.groupby('Geography Name')['Absorption Rate'].shift(1)
df_custom['4 Quarters Ago Absorption Rate']   =  df_custom.groupby('Geography Name')['Absorption Rate'].shift(4)

df_custom['QoQ Absorption Growth']        = round((df_custom['Absorption Rate'] - df_custom['Previous Quarter Absorption Rate']) * 100,0)
df_custom['YoY Absorption Growth']        = round((df_custom['Absorption Rate'] - df_custom['4 Quarters Ago Absorption Rate'])   * 100,0)

#Sales Volume
df_custom['Previous Quarter Total Sales Volume'] = df_custom.groupby('Geography Name')['Total Sales Volume'].shift(1)
df_custom['4 Quarters Ago Total Sales Volume']   = df_custom.groupby('Geography Name')['Total Sales Volume'].shift(4)

df_custom['QoQ Total Sales Volume Growth']        = round( (((df_custom['Total Sales Volume']  / df_custom['Previous Quarter Total Sales Volume']) - 1) * 100),                    0)
df_custom['YoY Total Sales Volume Growth']        = round( (((df_custom['Total Sales Volume']  / df_custom['4 Quarters Ago Total Sales Volume'])   - 1) * 100),                    0)

#Transactions
df_custom['Previous Quarter Sales Volume Transactions'] = df_custom.groupby('Geography Name')['Sales Volume Transactions'].shift(1)
df_custom['4 Quarters Ago Sales Volume Transactions']   = df_custom.groupby('Geography Name')['Sales Volume Transactions'].shift(4)

df_custom['QoQ Transactions Growth']         = round(  (((df_custom['Sales Volume Transactions']/df_custom['Previous Quarter Sales Volume Transactions']) - 1)  * 100)            ,0)
df_custom['YoY Transactions Growth']         = round(  (((df_custom['Sales Volume Transactions']/df_custom['4 Quarters Ago Sales Volume Transactions'])   - 1)  * 100)            ,0)


#market cap rate
df_custom['Market Cap Rate']                 = 0
df_custom['Previous Quarter Market Cap Rate'] = df_custom.groupby('Geography Name')['Market Cap Rate'].shift(1)
df_custom['4 Quarters Ago Market Cap Rate']   = df_custom.groupby('Geography Name')['Market Cap Rate'].shift(4)

df_custom['QoQ Market Cap Rate Growth']        = round((df_custom['Market Cap Rate'] - df_custom['Previous Quarter Market Cap Rate']) * 100,0)
df_custom['YoY Market Cap Rate Growth']        = round((df_custom['Market Cap Rate'] - df_custom['4 Quarters Ago Market Cap Rate'])   * 100,0)

# #Round  3 percentage variables we report in overview table
# df_custom['Market Cap Rate']            = round(df_custom['Market Cap Rate'],1)
df_custom['Vacancy Rate']               = round(df_custom['Vacancy Rate'],1)
df_custom['Absorption Rate']            = round(df_custom['Absorption Rate'],1)


print(df_custom)
df_custom.to_excel(clean_custom_file)






