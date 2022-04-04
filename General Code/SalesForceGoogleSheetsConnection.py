from io import BytesIO
import requests
import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from google.oauth2.service_account import Credentials
import os
import shapefile
from shapely.geometry import Point
from shapely.geometry.polygon import Polygon

#Specify file paths
dropbox_root                     = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)')               
msa_divisions_shapefile_location = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'Metro Divisions Shapefile', 'tl_2019_us_metdiv.shp')
msa_shapefile_location           = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'US CBSA Shapefile', 'cb_2018_us_cbsa_500k.shp')
county_shapefile_location        = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'US Counties Shapefile', 'cb_2018_us_county_500k.shp')


#Open the CoStar Markets export as dataframe
costar_markets_df                = pd.read_csv(os.path.join(dropbox_root, 'Research', 'Market Analysis', 'Market', 'CoStar Markets.csv'), dtype= {'CBSA Code': object})
costar_markets_df['CBSA Code']   = costar_markets_df['CBSA Code'].str[0:5]

#Open the MSA shapefile    
msa_map                          = shapefile.Reader(msa_shapefile_location)

#Open the MSA Division shapefile
msa_div_map                      = shapefile.Reader(msa_divisions_shapefile_location)

def FindMetroDivCode(point):
    #Loop through the MSA div map
    for i in range(len(msa_div_map) ):
        polygon               = Polygon(msa_div_map.shape(i).points)
        record                = msa_div_map.shapeRecord(i)
        div_code              = record.record['METDIVFP']
        if  polygon.contains(point):
            return(div_code)

def FindMSACode(point):
    #Loop through the MSA div map
    for i in range(len(msa_map) ):
        polygon               = Polygon(msa_map.shape(i).points)
        record                = msa_map.shapeRecord(i)
        msa_code              = record.record['CBSAFP']
        if  polygon.contains(point):
            return(msa_code)

def FindCountyFips(point):
    #Open the MSA shapefile
    county_map     = shapefile.Reader(county_shapefile_location)
    
    #Loop through the MSA div map
    for i in range(len(county_map) ):
        polygon               = Polygon(county_map.shape(i).points)
        record                = county_map.shapeRecord(i)
        state_fips            = record.record['STATEFP']
        county_fips           = record.record['COUNTYFP']
        county_fips           = state_fips + county_fips
        if  polygon.contains(point):
            return(county_fips)

def CountyNameFromFips(fips):
        master_county_list = pd.read_excel(os.path.join(dropbox_root, 'Research', 'Projects', 'Research Report Automation Project', 'Data', 'Area Reports Data', 'County_Master_List.xls'),
                dtype={'FIPS Code': object
                      }
                                          )
        master_county_list = master_county_list.loc[(master_county_list['FIPS Code'] == fips)]
        assert len(master_county_list) == 1
        county               = master_county_list['County Name'].iloc[0]
        return(county)

def FindCoStarMarket(msa_code, msa_div_code, sector):


    #Restrict to Markets (Metros)
    costar_df         = costar_markets_df.loc[(costar_markets_df['Analysis Type'] == 'Market') & (costar_markets_df['Property Type'] == sector) ]
    assert len(costar_df) > 1000    
    
    #Restrict to Market with matching CBSA code
    costar_msa_div_df = costar_df.loc[costar_df['CBSA Code'] == msa_div_code]

    
    # print(msa_code)
    # print(costar_df['Market Research Name','CBSA Code'].loc[costar_df['Market Research Name'] == 'CT - Stamford - Multifamily'])
    costar_msa_df     = costar_df.loc[costar_df['CBSA Code'] == msa_code]
    # print(costar_msa_df)
    
    #If the MSA Div code matches, use that, if not, try the MSA code
    if len(costar_msa_div_df) > 0:
        assert len(costar_msa_div_df) < 50
        market_name = costar_msa_div_df['Market Research Name'].iloc[-1]
        return(market_name)
    elif  len(costar_msa_div_df) == 0 and len(costar_msa_df) > 0:
        assert len(costar_msa_df) < 50
        market_name = costar_msa_df['Market Research Name'].iloc[-1]
        return(market_name)

def FormatGoogleSheetsURL(sheet_name):
    #Takes a google sheets name and reutrns a formated URL to pull the data from
    sheet_id   = '1fIP8dwH5hwSDMKEmOUdbnvwbwMZyAOClAm_4HDVVe5k'    
    url        = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return(url)

def GoogleSheetsURLToDF(url):
    #This function takes a google sheets url and returns a pandas dataframe with that data

    #Send request to fetch data
    r    = requests.get(url)
    data = r.content

    #Convert data into pandas dataframe
    df   = pd.read_csv(BytesIO(data), engine='python', keep_default_na=False)
    return(df)

def AuthorizeGoogle():
    global scope, credentials, gc
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive'
            ]
    project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects','Research Report Automation Project') 
    credentials = Credentials.from_service_account_file(os.path.join(project_location,'Code','General Code','GoogleCredentials.json'), scopes=scope)
    gc          = gspread.authorize(credentials)

def ProcessAreaReports(df):
    return(df)

def ProcessMarketReports(df):
    return(df)

def ProcessHoodReports(df):
    return(df)

def ProcessPropetyReports(df):
    df['Zip Code'] = df['Zip Code'].astype(str)
    
    df.loc[ (df['State'] == 'NY') & (df['County'] == 'New York'), 'Borough'] = 'Manhattan'
    df.loc[ (df['State'] == 'NY') & (df['County'] == 'Bronx'),    'Borough'] = 'Bronx'
    df.loc[ (df['State'] == 'NY') & (df['County'] == 'Kings'), 'Borough'] = 'Kings'
    df.loc[ (df['State'] == 'NY') & (df['County'] == 'Queens'), 'Borough'] = 'Queens'
    df.loc[ (df['State'] == 'NY') & (df['County'] == 'Richmond'), 'Borough'] = 'Staten Island'

    



    return(df)

def ProcessJobRecords(df):
    return(df)

#Authorize Google credentials and define our google sheet key
AuthorizeGoogle()
spreadsheet_key         = '1fIP8dwH5hwSDMKEmOUdbnvwbwMZyAOClAm_4HDVVe5k'

#Create our dataframes, one for each sheet
area_reports_df         = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Area Reports') )
market_reports_df       = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Market Reports') )
hood_reports_df         = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Hood Reports') )
property_records_df     = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'PROP') )
job_records_df          = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'JOB') )


#Create clean dataframes out of our raw dataframes 
area_reports_df_clean     = ProcessAreaReports(area_reports_df)
market_reports_df_clean   = ProcessMarketReports(market_reports_df)
hood_reports_df_clean     = ProcessHoodReports(hood_reports_df)
property_records_df_clean = ProcessPropetyReports(property_records_df)
job_records_df_clean      = ProcessJobRecords(job_records_df)


#Merge JOB and PROP dfs
merged_prop_job_df      = pd.merge(property_records_df_clean, job_records_df_clean, on=['Property Number'], how = 'left', ) 
assert len(merged_prop_job_df) == len(property_records_df_clean)


#Iterate through each row of the merged job x prop df and assign each record a market
for i in range(len(merged_prop_job_df)):

    lat, lon                 = merged_prop_job_df['Property Latitude'].iloc[i], merged_prop_job_df['Property Longitude'].iloc[i]
    property_coordinates     = Point(lon, lat) #lon, lat
    msa_div_code             = FindMetroDivCode(point=property_coordinates) 
    msa_code                 = FindMSACode(point=property_coordinates) 
    mf_costar_market         = FindCoStarMarket(msa_code = msa_code, msa_div_code = msa_div_code, sector = 'Multifamily')
    retail_costar_market     = FindCoStarMarket(msa_code = msa_code, msa_div_code = msa_div_code, sector = 'Retail')
    office_costar_market     = FindCoStarMarket(msa_code = msa_code, msa_div_code = msa_div_code, sector = 'Office')
    industrial_costar_market = FindCoStarMarket(msa_code = msa_code, msa_div_code = msa_div_code, sector = 'Industrial')
    prop_type                = merged_prop_job_df['Property Type'].iloc[i]
    
    if prop_type == ('Multifamily' or 'Mixed Use (Residential and Commercial)' or 'Mixed Use (Residential and Commercial) - Mixed-Use Residential Rental' or 'Multifamily - Cooperative Building'):
        costar_market = mf_costar_market
    elif prop_type == ('Industrial' or 'Industrial - Flex Space' or 'Industrial - Warehouse'):
        costar_market = industrial_costar_market
    elif prop_type == ('Office - Medical' or 'Office' or 'Office - Condominium Building'):
        costar_market = office_costar_market
    elif prop_type == ('Retail - Restaurant' or 'Retail' or 'Retail - Other Commercial & Retail' or 'Retail - Day Care Facility' or 'Retail - Convenience / Strip Center' or 'Retail - Automotive - Parking Structure'):
        costar_market = retail_costar_market
    else:
        costar_market = mf_costar_market


    merged_prop_job_df['Market'].iloc[i] = costar_market

#Merge the market back into the property df
merged_prop_job_df = merged_prop_job_df.fillna('')
merged_prop_job_df = merged_prop_job_df[['Property Number','Market']]
property_records_df_clean = property_records_df_clean.drop(columns='Market')
property_records_df_clean = pd.merge(property_records_df_clean, merged_prop_job_df, on=['Property Number'], how = 'left', ) 
property_records_df_clean = property_records_df_clean.fillna('')

#Re-order columns
column_names = ['Property Number', 'City',	'State',	'Zip Code',	'Market',	'Neighborhood/District',	'Borough',	'County', 'FIPS',	'CBSA']
property_records_df_clean = property_records_df_clean.reindex(columns=column_names)


#Upload our cleaned dataframes back to the google sheets
d2g.upload(df = property_records_df_clean, gfile = spreadsheet_key, wks_name = 'PROP TEST', row_names=False,)





# d2g.upload(df = area_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Area Reports TEST', row_names=False)
# d2g.upload(df = market_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Market Reports TEST', row_names=False)
# d2g.upload(df = hood_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Hood Reports TEST', row_names=False)
# d2g.upload(df = merged_prop_job_df, gfile = spreadsheet_key, wks_name = 'JOB PROP MERGE TEST', row_names=False,)
