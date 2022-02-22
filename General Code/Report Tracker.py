#Date: 02/02/2022
#Author: Mike Leahy
#Summary: Uses our 3 report csv summary files and produces a summary of jobs produced by each member of the research team

import os
import pandas as pd

#Define file paths
dropbox_root                        =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
main_output_location                =  os.path.join(dropbox_root,'Research','Market Analysis') 

#Read in our 3 csv reports as dataframes
area_report_df                      = pd.read_csv(os.path.join(main_output_location,'Area','Dropbox Areas.csv'),        encoding='latin-1')
market_report_df                    = pd.read_csv(os.path.join(main_output_location,'Market','CoStar Markets.csv'),     encoding='latin-1')
hood_report_df                      = pd.read_csv(os.path.join(main_output_location,'Neighborhood','Dropbox Neighborhoods.csv'),encoding='latin-1')

#Restrict to Final Reports
area_report_df                       = area_report_df.loc[area_report_df['Status'] == 'Final']
market_report_df                     = market_report_df.loc[market_report_df['Status'] == 'Final']
hood_report_df                       = hood_report_df.loc[hood_report_df['Status'] == 'Final']


#Collapse down each dataframe to the total done by each team member
area_report_df['Total Reports']      = 1
market_report_df['Total Reports']    = 1
hood_report_df['Total Reports']      = 1

area_report_df                      = area_report_df.groupby(['Version']).agg({'Total Reports': 'sum'}).reset_index()
market_report_df                    = market_report_df.groupby(['Version']).agg({'Total Reports': 'sum'}).reset_index()
hood_report_df                      = hood_report_df.groupby(['Version']).agg({'Total Reports': 'sum'}).reset_index()

area_report_df['Type']              = 'Area'
market_report_df['Type']            = 'Market'
hood_report_df['Type']              = 'Hood'


#Append the 3 dataframes together
kpi_df                              = pd.concat([area_report_df,market_report_df,hood_report_df])
kpi_df                              = kpi_df.sort_values(by=['Type','Version'])

print(kpi_df)
#Export as csv file
kpi_df.to_csv(os.path.join(dropbox_root,'Research','Projects','KPI','KPI.csv'),index=False)



