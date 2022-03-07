from io import BytesIO
import requests
import pandas as pd
import gspread
from df2gspread import df2gspread as d2g
from google.oauth2.service_account import Credentials
import os

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
    df   = pd.read_csv(BytesIO(data), engine='python')
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

#Authorize Google credentials and define our google sheet key
AuthorizeGoogle()
spreadsheet_key         = '1fIP8dwH5hwSDMKEmOUdbnvwbwMZyAOClAm_4HDVVe5k'

#Create 3 dataframes, one for each report type
area_reports_df         = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Area Reports') )
market_reports_df       = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Market Reports') )
hood_reports_df         = GoogleSheetsURLToDF(url = FormatGoogleSheetsURL(sheet_name = 'Hood Reports') )

#Create clean dataframes out of our raw dataframes 
area_reports_df_clean   = ProcessAreaReports(area_reports_df)
market_reports_df_clean = ProcessMarketReports(market_reports_df)
hood_reports_df_clean   = ProcessHoodReports(hood_reports_df)

#Upload our cleaned dataframes back to the google sheets
d2g.upload(df = area_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Area Reports TEST', row_names=True)
d2g.upload(df = market_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Market Reports TEST', row_names=True)
d2g.upload(df = hood_reports_df_clean, gfile = spreadsheet_key, wks_name = 'Hood Reports TEST', row_names=True)
