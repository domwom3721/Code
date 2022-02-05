#Date: 02/02/2022
#Author: Mike Leahy
#Summary: Uses our 3 report csv summary files and produces a summary of jobs produced by each member of the research team

import os
import pandas as pd

#Define file paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
main_output_location           =  os.path.join(dropbox_root,'Research','Market Analysis') 

#Read in our 3 csv reports as dataframes
area_report_df    = pd.read_csv(os.path.join(main_output_location,'Area','Dropbox Areas.csv'),        encoding='latin-1')
market_report_df  = pd.read_csv(os.path.join(main_output_location,'Market','CoStar Markets.csv'),     encoding='latin-1')
hood_report_df    = pd.read_csv(os.path.join(main_output_location,'Neighborhood','Dropbox Neighborhoods.csv'),encoding='latin-1')

#Restrict to Final Reports
area_report_df    = area_report_df.loc[area_report_df['Status'] == 'Final']
market_report_df  = market_report_df.loc[market_report_df['Status'] == 'Final']
hood_report_df    = hood_report_df.loc[hood_report_df['Status'] == 'Final']


#Collapse down each dataframe to the total done by each team member
area_report_df['Total Area Reports']      = 1
market_report_df['Total Market Reports']  = 1
hood_report_df['Total Hood Reports']      = 1

area_report_df                      = area_report_df.groupby(['Assigned To','Version']).agg({'Total Area Reports': 'sum'}).reset_index()
market_report_df                    = market_report_df.groupby(['Assigned To','Version']).agg({'Total Market Reports': 'sum'}).reset_index()
hood_report_df                      = hood_report_df.groupby(['Assigned To','Version']).agg({'Total Hood Reports': 'sum'}).reset_index()

#Restrict to latest version
latest_market_verion = '2021 Q4'
latest_area_verion   = '2021 Q4'
latest_hood_version  = 2022


area_report_df    = area_report_df.loc[area_report_df['Version']     == latest_area_verion ]
market_report_df  = market_report_df.loc[market_report_df['Version'] == latest_market_verion]
hood_report_df    = hood_report_df.loc[hood_report_df['Version']     == latest_hood_version]

#Rename verion variables
area_report_df                      = area_report_df.rename(columns={"Version": "Area Report Version"})
market_report_df                    = market_report_df.rename(columns={"Version": "Market Report Version"})
hood_report_df                      = hood_report_df.rename(columns={"Version": "Hood Report Version"})


#Merge the 3 dataframes together
kpi_df                              = pd.merge(area_report_df,market_report_df, on=['Assigned To'],how = 'left') 
kpi_df                              = pd.merge(kpi_df,hood_report_df, on=['Assigned To'],how = 'left') 


print(kpi_df)
kpi_df.to_csv(os.path.join(main_output_location,'KPI.csv'),index=False)