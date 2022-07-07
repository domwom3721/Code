#Cleans the raw data we download from CoStar each quarter and exports clean data into csv files
#By Research

#Import packages we will be using
import os
from numpy import int64
import pandas as pd

#Define file location pre paths
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project')  
costar_data_location           =  os.path.join(project_location,'Data','Market Reports Data','CoStar Data') 

#Define location of raw CoStar data files
raw_multifamily_file           =  os.path.join(costar_data_location,'Raw Data','mf.csv') 
raw_office_file                =  os.path.join(costar_data_location,'Raw Data','office.csv') 
raw_retail_file                =  os.path.join(costar_data_location,'Raw Data','retail.csv') 
raw_industrial_file            =  os.path.join(costar_data_location,'Raw Data','industrial.csv') 

raw_multifamily_slices_file    =  os.path.join(costar_data_location,'Raw Data','mf_slices.csv') 
raw_office_slices_file         =  os.path.join(costar_data_location,'Raw Data','office_slices.xlsx') 
raw_retail_slices_file         =  os.path.join(costar_data_location,'Raw Data','retail_slices.xlsx') 
raw_industrial_slices_file     =  os.path.join(costar_data_location,'Raw Data','industrial_slices.csv') 


#Import raw CoStar data as pandas dataframes
df_multifamily  = pd.read_csv(raw_multifamily_file,
                dtype={'Sales Volume Transactions': object
                      }      ) 

df_office       = pd.read_csv(raw_office_file,
                  dtype={'Sales Volume Transactions':   object,
                         'Total Sales Volume':          object,
                         'Sold Building SF':            object,
                        'Transaction Sale Price/SF':    object,
                        'Under Construction Buildings': object,
                        'Asset Value':                  object,
                        'Availability Rate':            object,
                        'Available SF':                 object,
                        'Average Sale Price':           object,
                        'Existing Buildings':           object,
                        'Asset Value':                  object,
                        'Total Sales Volume':           object,
                        'Cap Rate':                     object,
                        'Existing Buildings':           object,
                        'Market Sale Price Growth':     object,
                        'Occupancy Rate':               object,
                        'Median Cap Rate':              object,
                        'Under Construction Buildings': object,
                        'Year':                         object,
                        'Sales Volume Transactions':    object,
                        'Net Absorption SF':            object,
                        'Inventory SF':                 object,
                        'Under Construction SF':        object,
                        'Market Rent/SF':               object,
                        'Market Rent Growth':           object,
                        'Market Rent Growth 12 Mo':     object,
                        'Available SF':                 object,
                        'Vacancy Rate':                 object,
                        'Median Cap Rate':              object,
                        'Median Price/Bldg SF':         object,
                        'Demand SF':                    object,

                        }     
                            )

df_retail       = pd.read_csv(raw_retail_file,
                  dtype={'Sales Volume Transactions':  object,
                       'Cap Rate Transactions'    :    object,
                       'Gross Delivered Buildings':    object,
                       'Sold Building SF':             object,
                       'Total Sales Volume':           object,
                       'Office Gross Rent Sublet':     object,
                       'Office Gross Rent Direct':     object,
                       'Office Gross Rent Overall':    object,
                       'Transaction Sale Price/SF':    object,
                       'Under Construction Buildings': object
                       }
                            )

df_industrial   = pd.read_csv(raw_industrial_file,
                  dtype={'Sales Volume Transactions':    object,
                         'Sold Building SF':             object,
                         'Total Sales Volume':           object,
                         'Transaction Sale Price/SF':    object,
                         'Under Construction Buildings': object,
                         'Vacancy Rate':                 float
                        }
                             )  		

print('Importing Raw CoStar Data')

#Import the raw slices data from Costar where the markets are broken down by the quality or type of the properties
if os.path.exists(raw_multifamily_slices_file):
    df_multifamily_slices  = pd.read_csv(raw_multifamily_slices_file,
                    dtype={'Sales Volume Transactions': object,
                          'Market Effective Rent/Unit': float,
                          'Inventory Units': int64
                        }      
                                        ) 
    
if os.path.exists(raw_office_slices_file):
    df_office_slices       = pd.read_excel(raw_office_slices_file,
                    dtype={'Sales Volume Transactions':   object,
                            'Total Sales Volume':           object,
                            'Transaction Sale Price/SF':    object,
                            'Under Construction Buildings': int64,
                            'Vacancy Rate':                 float,
                            'Inventory SF':                 int64,
                            }     
                                        )

if os.path.exists(raw_retail_slices_file):
    df_retail_slices       = pd.read_excel(raw_retail_slices_file,
                    dtype={'Sales Volume Transactions':  object,
                        'Cap Rate Transactions':        object,
                        'Gross Delivered Buildings':    object,
                        'Sold Building SF':             object,
                        'Total Sales Volume':           object,
                        'Office Gross Rent Sublet':     object,
                        'Office Gross Rent Direct':     object,
                        'Office Gross Rent Overall':    object,
                        'Transaction Sale Price/SF':    object,
                        'Under Construction Buildings': object,
                        'Inventory SF':                 int64,
                        }
                                        )

if os.path.exists(raw_industrial_slices_file):
    df_industrial_slices   = pd.read_csv(raw_industrial_slices_file,
                    dtype={'Sales Volume Transactions':    object,
                            'Sold Building SF':             object,
                            'Total Sales Volume':           object,
                            'Transaction Sale Price/SF':    object,
                            'Under Construction Buildings': object,
                            'Vacancy Rate':                 float,
                            'Inventory SF':                 int64,
                            }
                                        )

print('Importing Raw CoStar Sliced Data')

#Define data cleaning functions
def DropClusters(df): 
    #Drops rows that report data on the cluster geography type
    df = df.loc[df['Geography Type'] != 'Cluster']
    return(df)
print('Dropping Clusters')

def DropUrban(df): 
    #Drops rows that report data on the urban geography type
    df = df.loc[df['Geography Type'] != 'Location Type:Urban']
    return(df)

print('Dropping Urban')


def DropSuburban(df): 
    #Drops rows that report data on the suburban geography type
    df = df.loc[df['Geography Type'] != 'Location Type:Suburban']
    return(df)

print('Dropping Suburban')

def DropCBD(df): 
    #Drops rows that report data on the cbd geography type
    df = df.loc[df['Geography Type'] != 'Location Type:CBD']
    return(df)

print('Dropping CBD')

def KeepLast10Years(df,groupbylist): 
    #Cut down to last 10 years
    df = df.groupby(groupbylist).tail(40)
    return(df)

print('Keeping Last 10 Years')

def SortData(df):
     #Sorts by geography and quarter
    df = df.sort_values(by=['Geography Name','Period'])
    return(df)

print('Sorting Data')

def CreateYearAndQuarterVariables(df): 
    #Seperates the period variable into 2 components (year and quarter)
    df.loc[:,'Year']           =   df.loc[:,'Period'].str[:4]
    df.loc[:,'Quarter']        =   df.loc[:,'Period'].str[5:]
    return(df)

print('Creating Year & Quarter Variables')

def DropExtraVariables(df,sector): 
    #Drops the variables we don't use in our analysis
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

print('Dropping Extra Variables')

def DestringVariablesConvertToNumeric(df,sector):

    if sector == 'Multifamily':
        vars_list_to_destring = [
                'Average Sale Price',
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
        vars_list_to_destring = [
                'Average Sale Price',
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
                'Net Absorption SF',
                'Inventory SF',
                'Under Construction SF',
                'Market Rent/SF',
                'Market Rent Growth',
                'Market Rent Growth 12 Mo',
                'Available SF'
                ]


    for var in vars_list_to_destring:
        #only do the following for string variables
        if df[var].dtype == 'object': 
            df[var] = df[var].str.replace('$', '', regex=False)
            df[var] = df[var].str.replace(',', '', regex=False)
            df[var] = df[var].str.replace('%', '', regex=False)
            df[var] = pd.to_numeric(df[var])
    return(df)

def FillBlanksWithZero(df,sector):
    if sector == 'Multifamily':
        var_list_to_replace_blanks = ['Sales Volume Transactions','Total Sales Volume','Under Construction Units','Absorption Units']
    else:
        var_list_to_replace_blanks = ['Sales Volume Transactions','Total Sales Volume','Under Construction SF','Net Absorption SF','Sold Building SF']

    for var in var_list_to_replace_blanks:
        df[var] = df[var].fillna(0)

    return(df)

def ConvertPercenttoPercentagePoints(df,sector):
    if sector == 'Multifamily':
        rate_vars = ['Absorption %',
                    'Vacancy Rate',
                    'Cap Rate',
                    'Market Asking Rent Growth',
                    'Market Asking Rent Growth 12 Mo',
                    'Market Cap Rate',
                    'Market Effective Rent Growth',
                    'Market Effective Rent Growth 12 Mo',
                    'Market Sale Price Growth',
                    'Median Cap Rate',
                    'Occupancy Rate',
                    ]
    else:
        rate_vars = ['Vacancy Rate',
                    'Availability Rate',
                    'Market Cap Rate', ]
                   #'Market Rent Growth', 
                   #'Market Rent Growth 12 Mo' 
    
    for var in rate_vars:
            df[var] = round(df[var] * 100,4)
    return(df)

def MainClean(df,sector): #Calls all cleaning functions and returns cleaned dataframes
    df = SortData(df)
    df = CreateYearAndQuarterVariables(df)
    df = DropExtraVariables(df,sector)
    df = DestringVariablesConvertToNumeric(df,sector)
    df = ConvertPercenttoPercentagePoints(df,sector)
    df = FillBlanksWithZero(df,sector)
    return(df)

def MainCleanSlices(df,sector): #Calls cleaning functions and returns cleaned dataframes for our sliced data
    if sector == 'Multifamily':
        df = df[['Property Class Name','Period','Slice','As Of','Geography Name','Property Type','Vacancy Rate','Market Effective Rent/Unit','Market Effective Rent Growth','Market Effective Rent Growth 12 Mo','Inventory Units']]
    else:
        df = df[['Property Class Name','Period','Slice','As Of','Geography Name','Property Type','Vacancy Rate','Market Rent/SF','Market Rent Growth','Market Rent Growth 12 Mo','Inventory SF','Under Construction SF','Net Delivered SF 12 Mo','Availability Rate','Net Absorption SF 12 Mo']]
        df['Availability Rate'] = df['Availability Rate'] * 100

        #Replace NA with 0
        df['Under Construction SF'] = df['Under Construction SF'].fillna(0)
    
    #Clean the vacancy rate variable by removing the percentage sign
    if df['Vacancy Rate'].dtype == 'object':
        df['Vacancy Rate'] = df['Vacancy Rate'].str.replace('%','')
    else:
        df['Vacancy Rate'] = round(df['Vacancy Rate'] * 100,2)

    #Clean the rent variables by removing the dollar sign and commas
    if sector == 'Multifamily':
        if df['Market Effective Rent/Unit'].dtype == 'object':
            df['Market Effective Rent/Unit'] = df['Market Effective Rent/Unit'].str.replace('$','')
            df['Market Effective Rent/Unit'] = df['Market Effective Rent/Unit'].str.replace(',','',5)
            df['Market Effective Rent/Unit'] = df['Market Effective Rent/Unit'].astype(float)
        

    else:
        if df['Vacancy Rate'].dtype == 'object':
            df['Market Rent/SF'] = df['Market Rent/SF'].str.replace('$','')
            df['Market Rent/SF'] = df['Market Rent/SF'].str.replace(',','',5)
            df['Market Rent/SF'] = df['Market Rent/SF'].astype(float)
        
        if df['Inventory SF'].dtype == 'object':
            df['Inventory SF']             = df['Inventory SF'].str.replace(',','')
            df['Inventory SF']             = df['Inventory SF'].astype(int64)
            
    #Remove "Center" from slice name
    df['Slice']                = df['Slice'].str.replace(' Center', '', regex=False)
    #df['Geography Name']       = df['Geography Name'].str.replace('New York City', 'Manhattan', regex=False)
    
    #Drop the aggregate slice
    df                         =   df.loc[df['Slice'] != 'All']
    df.loc[:,'Year']           =   df.loc[:,'Period'].str[:4]
    df.loc[:,'Quarter']        =   df.loc[:,'Period'].str[5:]

    return(df)

#Pass our 4 dataframes into our main cleaning function which calls all the other cleaning functions
df_multifamily          =  MainClean(df_multifamily,'Multifamily')
df_office               =  MainClean(df_office,'Office')
df_retail               =  MainClean(df_retail,'Retail')
df_industrial           =  MainClean(df_industrial,'Industrial')

df_multifamily_slices   =  MainCleanSlices(df_multifamily_slices,'Multifamily')
df_office_slices        =  MainCleanSlices(df_office_slices,'Office')
df_retail_slices        =  MainCleanSlices(df_retail_slices,'Retail')
df_industrial_slices    =  MainCleanSlices(df_industrial_slices,'Industrial')

print('Cleaning Sliced Data')

#Loop through the 4 dataframes: create variables we will use in our report/figures 
for df in [df_multifamily,df_office,df_retail,df_industrial]:

    #df['Geography Name'] = df['Geography Name'].str.replace('New York City', 'Manhattan', regex=False)

    #Clean the sqft and unit variables seperately
    if df.equals(df_multifamily):

        #Create laggd variables
        df['Lagged Inventory Units']                      = df.groupby('Geography Name')['Inventory Units'].shift(1)
      
        #Create variable for apt absorption rate
        df['Absorption Rate']                             = round(((df['Absorption Units']/df['Inventory Units']) * 100), 2) 

        #Create variable for inventory growth rates
        df['Inventory Units Growth']                      = df['Inventory Units'] - df['Lagged Inventory Units']
        df['Inventory Growth']                            = round(((df['Inventory Units'] / df['Lagged Inventory Units']) - 1)  * 100,2)

        #Create variable for percent under construction
        df['Under Construction %']                        = (df['Under Construction Units']/df['Inventory Units'] ) *100

        #Asset Value per unit
        df['Asset Value/Unit']                            = round((df['Asset Value']/df['Inventory Units']),2)
        df['Previous Quarter Asset Value/Unit']           = df.groupby('Geography Name')['Asset Value/Unit'].shift(1)
        df['4 Quarters Ago Asset Value/Unit']             = df.groupby('Geography Name')['Asset Value/Unit'].shift(4)

        df['QoQ Asset Value/Unit Growth']                 = round( (((df['Asset Value/Unit']  / df['Previous Quarter Asset Value/Unit']) - 1) * 100),                    1)
        df['YoY Asset Value/Unit Growth']                 = round( (((df['Asset Value/Unit']  / df['4 Quarters Ago Asset Value/Unit'])   - 1) * 100),                    1)

        #Market Rent
        df['Previous Quarter Market Effective Rent/Unit'] = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(1)
        df['4 Quarters Ago Market Effective Rent/Unit']   = df.groupby('Geography Name')['Market Effective Rent/Unit'].shift(4)

        df['QoQ Market Effective Rent/Unit Growth']       = round( (((df['Market Effective Rent/Unit']   / df['Previous Quarter Market Effective Rent/Unit']) - 1) * 100),                    1)
        df['YoY Market Effective Rent/Unit Growth']       = round( (((df['Market Effective Rent/Unit']  / df['4 Quarters Ago Market Effective Rent/Unit'])   - 1) * 100),                    1)
        
        #Absorption Units
        df['Previous Quarter Absorption Units']           = df.groupby('Geography Name')['Absorption Units'].shift(1)
        df['4 Quarters Ago Absorption Units']             = df.groupby('Geography Name')['Absorption Units'].shift(4)

        df['QoQ Absorption Units Growth']                 = round((df['Absorption Units']   - df['Previous Quarter Absorption Units'])    / abs(df['Previous Quarter Absorption Units'])  * 100,1)              
        df['YoY Absorption Units Growth']                 = round((df['Absorption Units']   - df['4 Quarters Ago Absorption Units'])      /  abs(df['4 Quarters Ago Absorption Units'])   * 100 ,1)           
           

    else:            
        #Create laggd variables
        df['Lagged Inventory SF']                         = df.groupby('Geography Name')['Inventory SF'].shift(1)

        #Create variable for absorption rate 
        df['Absorption Rate']                            = round((df['Net Absorption SF'] / df['Inventory SF']) * 100,2 )

        #Absorption SF
        df['Previous Quarter Net Absorption SF']         = df.groupby('Geography Name')['Net Absorption SF'].shift(1)
        df['4 Quarters Ago Net Absorption SF']           = df.groupby('Geography Name')['Net Absorption SF'].shift(4)

        df['QoQ Net Absorption SF Growth']               = round((df['Net Absorption SF']   - df['Previous Quarter Net Absorption SF'])    / abs(df['Previous Quarter Net Absorption SF'])  * 100,1)              
        df['YoY Net Absorption SF Growth']               = round((df['Net Absorption SF']   - df['4 Quarters Ago Net Absorption SF'])      /  abs(df['4 Quarters Ago Net Absorption SF'])   * 100 ,1)           
           
        #Availability Rate 
        df['Previous Quarter Availability Rate']         = df.groupby('Geography Name')['Availability Rate'].shift(1)
        df['4 Quarters Ago Availability Rate']           = df.groupby('Geography Name')['Availability Rate'].shift(4)
    
        df['QoQ Availability Rate Growth']               = round((df['Availability Rate'] - df['Previous Quarter Availability Rate']) * 100,0)
        df['YoY Availability Rate Growth']               = round((df['Availability Rate'] - df['4 Quarters Ago Availability Rate'])   * 100,0)

        df['Availability Rate']                          = round(df['Availability Rate'],1)

        #Market Rent
        df['Previous Quarter Market Rent/SF']            = df.groupby('Geography Name')['Market Rent/SF'].shift(1)
        df['4 Quarters Ago Market Rent/SF']              = df.groupby('Geography Name')['Market Rent/SF'].shift(4)

        df['QoQ Rent Growth']                            = round( (((df['Market Rent/SF']  / df['Previous Quarter Market Rent/SF']) - 1) * 100),                    1)
        df['YoY Rent Growth']                            = round( (((df['Market Rent/SF']  / df['4 Quarters Ago Market Rent/SF'])   - 1) * 100),                    1)

        #Create variable for inventory growth rate
        df['Inventory SF Growth']                        = df['Inventory SF'] - df['Lagged Inventory SF']  
        df['Inventory Growth']                           = round(((df['Inventory SF'] / df['Lagged Inventory SF']) - 1)  * 100,2)

        #Create variable for percent under construction
        df['Under Construction %']                       = (df['Under Construction SF']/df['Inventory SF'] ) *100

        #Asset Value per sqft
        df['Asset Value/Sqft']                           = round((df['Asset Value']/df['Inventory SF']),2)

        df['Previous Quarter Asset Value/Sqft']          = df.groupby('Geography Name')['Asset Value/Sqft'].shift(1)
        df['4 Quarters Ago Asset Value/Sqft']            = df.groupby('Geography Name')['Asset Value/Sqft'].shift(4)

        df['QoQ Asset Value/Sqft Growth']                = round( (((df['Asset Value/Sqft']  / df['Previous Quarter Asset Value/Sqft']) - 1) * 100),                    1)
        df['YoY Asset Value/Sqft Growth']                = round( (((df['Asset Value/Sqft']  / df['4 Quarters Ago Asset Value/Sqft'])   - 1) * 100),                    1)


    #Making Variables for all sectors
    df['Previous Quarter Vacancy']                       = df.groupby('Geography Name')['Vacancy Rate'].shift(1)
    df['4 Quarters Ago Vacancy']                         = df.groupby('Geography Name')['Vacancy Rate'].shift(4)
    
    df['QoQ Vacancy Growth']                             = round((df['Vacancy Rate'] - df['Previous Quarter Vacancy']) * 100,0)
    df['YoY Vacancy Growth']                             = round((df['Vacancy Rate'] - df['4 Quarters Ago Vacancy'])   * 100,0)

    #Absorption
    df['Previous Quarter Absorption Rate']               =  df.groupby('Geography Name')['Absorption Rate'].shift(1)
    df['4 Quarters Ago Absorption Rate']                 =  df.groupby('Geography Name')['Absorption Rate'].shift(4)

    df['QoQ Absorption Growth']                          = round((df['Absorption Rate'] - df['Previous Quarter Absorption Rate']) * 100,0)
    df['YoY Absorption Growth']                          = round((df['Absorption Rate'] - df['4 Quarters Ago Absorption Rate'])   * 100,0)
    
    #Sales Volume
    df['Previous Quarter Total Sales Volume']            = df.groupby('Geography Name')['Total Sales Volume'].shift(1)
    df['4 Quarters Ago Total Sales Volume']              = df.groupby('Geography Name')['Total Sales Volume'].shift(4)
    
    df['QoQ Total Sales Volume Growth']                  = round( (((df['Total Sales Volume']  / df['Previous Quarter Total Sales Volume']) - 1) * 100),                    0)
    df['YoY Total Sales Volume Growth']                  = round( (((df['Total Sales Volume']  / df['4 Quarters Ago Total Sales Volume'])   - 1) * 100),                    0)

    #Transactions
    df['Previous Quarter Sales Volume Transactions']     = df.groupby('Geography Name')['Sales Volume Transactions'].shift(1)
    df['4 Quarters Ago Sales Volume Transactions']       = df.groupby('Geography Name')['Sales Volume Transactions'].shift(4)
    
    df['QoQ Transactions Growth']                        = round(  (((df['Sales Volume Transactions']/df['Previous Quarter Sales Volume Transactions']) - 1)  * 100)            ,0)
    df['YoY Transactions Growth']                        = round(  (((df['Sales Volume Transactions']/df['4 Quarters Ago Sales Volume Transactions'])   - 1)  * 100)            ,0)

    #Market cap rate
    df['Previous Quarter Market Cap Rate']               = df.groupby('Geography Name')['Market Cap Rate'].shift(1)
    df['4 Quarters Ago Market Cap Rate']                 = df.groupby('Geography Name')['Market Cap Rate'].shift(4)
    
    df['QoQ Market Cap Rate Growth']                     = round((df['Market Cap Rate'] - df['Previous Quarter Market Cap Rate']) * 100,0)
    df['YoY Market Cap Rate Growth']                     = round((df['Market Cap Rate'] - df['4 Quarters Ago Market Cap Rate'])   * 100,0)

    #Round 3 percentage variables we report in overview table
    df['Market Cap Rate']                                = round(df['Market Cap Rate'],2)
    df['Vacancy Rate']                                   = round(df['Vacancy Rate'],2)
    df['Absorption Rate']                                = round(df['Absorption Rate'],2)

print('Creating New Variables')

#Keep last 10 years only 
df_multifamily          =  KeepLast10Years(df_multifamily,groupbylist= ['Geography Name'])
df_office               =  KeepLast10Years(df_office,groupbylist= ['Geography Name'])
df_retail               =  KeepLast10Years(df_retail,groupbylist= ['Geography Name'])
df_industrial           =  KeepLast10Years(df_industrial,groupbylist= ['Geography Name'])

df_multifamily_slices   =  KeepLast10Years(df_multifamily_slices,groupbylist= ['Geography Name','Slice'])
df_office_slices        =  KeepLast10Years(df_office_slices,groupbylist= ['Geography Name','Slice'])
df_retail_slices        =  KeepLast10Years(df_retail_slices,groupbylist= ['Geography Name','Slice'])
df_industrial_slices    =  KeepLast10Years(df_industrial_slices,groupbylist= ['Geography Name','Slice'])

print('Keeping 10 Years of Data')

#Export Cleaned Data Files
df_multifamily.to_csv(os.path.join(costar_data_location,'Clean Data','mf_clean.csv'),index=False)
df_office.to_csv(os.path.join(costar_data_location, 'Clean Data','office_clean.csv'),index=False)
df_retail.to_csv(os.path.join(costar_data_location,'Clean Data','retail_clean.csv',),index=False)
df_industrial.to_csv(os.path.join(costar_data_location,'Clean Data','industrial_clean.csv'),index=False)

df_multifamily_slices.to_csv(os.path.join(costar_data_location,'Clean Data','mf_slices_clean.csv'),index=False)
df_office_slices.to_csv(os.path.join(costar_data_location,'Clean Data','office_slices_clean.csv'),index=False)
df_retail_slices.to_csv(os.path.join(costar_data_location,'Clean Data','retail_slices_clean.csv',),index=False)
df_industrial_slices.to_csv(os.path.join(costar_data_location,'Clean Data','industrial_slices_clean.csv'),index=False)

print('Exporting Excel Files')

print('Cleaning Complete')