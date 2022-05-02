#cleans raw data downloads from realtor.com each month and exports clean data into csv files
#Dom 8/3/2021
#packages needed
import os
import pandas as pd
import numpy as np

#define file paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')
realtor_data_location          =  os.path.join(project_location,'Data','Realtor Data')

#Define location of raw Realtor data files
raw_national_file              =  os.path.join(realtor_data_location,'national.csv') 
raw_state_file                 =  os.path.join(realtor_data_location,'state.csv') 
raw_metro_file                 =  os.path.join(realtor_data_location,'metro.csv') 
raw_county_file                =  os.path.join(realtor_data_location,'county.csv')
raw_zip_file                   =  os.path.join(realtor_data_location,'zip.csv')

#import raw realtor.com data as pandas data frames
df_national_residential = pd.read_csv(raw_national_file)


df_state_residential = pd.read_csv(raw_state_file)


df_metro_residential = pd.read_csv(raw_metro_file)


df_county_residential = pd.read_csv(raw_county_file,dtype={'county_fips': object}  
                                    
                                    )

df_zip_residential = pd.read_csv(raw_zip_file)


#cleaning data
def SortData(df): #Sorts by geography and quarter
    if level == 'state':
        df = df.sort_values(by=['state_id','month_date_yyyymm'])

    elif level == 'metro':
        df = df.sort_values(by=['cbsa_code','month_date_yyyymm'])
    
    elif level == 'zip':
        df = df.sort_values(by=['postal_code','month_date_yyyymm'])
    
    elif level == 'national':
        df = df.sort_values(by=['month_date_yyyymm'])
    
    elif level == 'county':
        df = df.sort_values(by=['county_fips','month_date_yyyymm'])

    return(df)

def FillInFips(df):
    df['county_fips'] = df['county_fips'].str.zfill(5)
    return(df)

def CreateYearAndMonthVariables(df): #seperates the month_date_yyyymm variable into 2 components (year and month)
    df.loc[:,'Year']           =   df.loc[:,'month_date_yyyymm'].str[:4]
    df.loc[:,'Month']        =   df.loc[:,'month_date_yyyymm'].str[5:]
    return(df)

def KeepLast4Years(df,groupbylist): #Cut down to last 4 years
    df = df.groupby(groupbylist).tail(48)
    return(df)

def ConvertPercenttoPercentagePoints(df):
    
    rate_vars = ['median_listing_price_mm',
                'median_listing_price_yy',
                'active_listing_count_mm',
                'active_listing_count_yy',
                'median_days_on_market_mm',
                'median_days_on_market_yy',
                'new_listing_count_mm',
                'new_listing_count_yy',
                'price_increased_count_mm',
                'price_increased_count_yy',
                'price_reduced_count_mm',
                'price_reduced_count_yy',
                'pending_listing_count_mm',
                'pending_listing_count_yy',
                'median_listing_price_per_square_foot_mm',
                'median_listing_price_per_square_foot_yy',
                'median_square_feet_mm',
                'median_square_feet_yy',
                'average_listing_price_mm',
                'average_listing_price_yy',
                'total_listing_count_mm',
                'total_listing_count_yy',
                'pending_ratio_mm',
                'pending_ratio_yy',
 
                ]

    
    for var in rate_vars:
       df[var] = round((df[var] * 100),2)
    return(df)

#Loop through the 5 dataframes: create variables we will use in our report/figures 
for df in [df_national_residential,df_state_residential,df_metro_residential,df_county_residential,df_zip_residential]:

    #df['Geography Name'] = df['Geography Name'].str.replace('New York City', 'Manhattan', regex=False)

        #Create Quarterly lagged variables since Annual and Month-over-month are provided
        df['Lagged Total Listing Count']       = df.groupby('Geography Name')['total_listing_count'].shift(3)
      

        #Create variable for absorption rate
        df['Absorption Rate'] = round(  ((df['pending_listing_count']/df['total_listing_count']) * 100)  ,2) 

        
        #Create variable for QoQ inventory growth rate
        df['Inventory Growth'] = round(((df['total_listing_count'] / df['Lagged Total Listing Count']) - 1)  * 100,2)

        #Create variable for percent under construction
        #df['Under Construction %'] = (df['Under Construction Units']/df['Inventory Units'] ) *100

        #Average Listing Price
        df['Lagged Avg Listing Price']       = df.groupby('Geography Name')['average_listing_price'].shift(1)
        df['Previous Quarter Avg Listing Price'] = df.groupby('Geography Name')['average_listing_price'].shift(3)
        df['4 Quarters Ago Avg Listing Price']   = df.groupby('Geography Name')['average_listing_price'].shift(12)
        df['Pandemic Listing Price']   = df.groupby('Geography Name')['average_listing_price'].shift(18)

        df['QoQ Avg Listing Price Growth']        = round( (((df['average_listing_price']  / df['Previous Quarter Avg Listing Price']) - 1) * 100),                    1)
        df['YoY Avg Listing Price Growth']        = round( (((df['average_listing_price']  / df['4 Quarters Ago Avg Listing Price'])   - 1) * 100),                    1)
        df['Pandemic Avg Listing Price Growth']   = round( (((df['average_listing_price']  / df['Pandemic Listing Price'])   - 1) * 100))))






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
           



#Create variable for absorption rate
#median_listing_price
#active_listing_count
#median_days_on_market
#new_listing_count
#price_increased_count
#price_reduced_count
#pending_listing_count
#median_listing_price_per_square_foot
#median_square_feet
#average_listing_price
#total_listing_count
#pending_ratio









def MainCleaningFunction(df):
    df = SortData(df)
    df = KeepLast4Years(df)
    df = CreateYearAndMonthVariables(df)
    df = ConvertPercenttoPercentagePoints(df)
    
    if level == 'county':
        df = FillInFips(df)

    return(df)
    
#Keep last 4 years only 
df_national_residential =  KeepLast4Years(df_national_residential,groupbylist= ['country'])
df_state_residential    =  KeepLast4Years(df_state_residential,groupbylist= ['state_id'])
df_metro_residential    =  KeepLast4Years(df_metro_residential,groupbylist= ['cbsa_code'])
df_county_residential   =  KeepLast4Years(df_county_residential,groupbylist= ['county_fips'])
df_zip_residential      =  KeepLast4Years(df_zip_residential,groupbylist= ['postal_code'])

#Main cleaning loop
for df,level in zip([df_national_residential,df_state_residential,df_metro_residential,df_county_residential,df_zip_residential],['national','state','metro','county','zip']):
    print(df,level)
    df = MainCleaningFunction(df)
    print(df)


df_national_residential.to_csv(os.path.join(realtor_data_location,'national_clean.csv'))
df_state_residential.to_csv(os.path.join(realtor_data_location,'state_clean.csv'))
df_metro_residential.to_csv(os.path.join(realtor_data_location,'metro_clean.csv'))
df_county_residential.to_csv(os.path.join(realtor_data_location,'county_clean.csv'))
df_zip_residential.to_csv(os.path.join(realtor_data_location,'zip_clean.csv'))
print('FINISHED')























