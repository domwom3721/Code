#Author: Mike Leahy
#Date:   06/22/2021
#Summary: This program identifies zip codes that are not within any Costar market/submarket, 
#         identify the county of these zips codes, and export a csv with this information

import pandas as pd
import os 

#Define file paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(dropbox_root,'Research Report Automation Project') #Main Folder that stores all output, code, and documentation

output_location                = os.path.join(project_location,'Output','Market Reports')         #The folder where we store our current reports, testing folder
data_location                  = os.path.join(project_location,'Data')
map_location                   = os.path.join(project_location,'Data','Maps','CoStar Maps')       #Folder with clean CoStar CSV files
costar_data_location           = os.path.join(project_location,'Data','Costar Data')              #Folders with maps png files


#Import Zip Code to CoStar Submarket crosswalk
df_zipcodes_submarket_crosswalk  = pd.read_excel(os.path.join(costar_data_location,'Zip to Submarket.xlsx'), dtype={'PostalCode': object} ) 
df_zipcodes_submarket_crosswalk.rename(columns={'PostalCode': 'ZIP'}, inplace=True) #Rename zip code variable

#cut down to main 4 property types
df_zipcodes_submarket_crosswalk = df_zipcodes_submarket_crosswalk.loc[(df_zipcodes_submarket_crosswalk['PropertyType'] == ('Multi-Family'))| 
                                                                      (df_zipcodes_submarket_crosswalk['PropertyType'] == ('Office'))      | 
                                                                      (df_zipcodes_submarket_crosswalk['PropertyType'] == ('Retail'))      |
                                                                      (df_zipcodes_submarket_crosswalk['PropertyType'] == ('Industrial'))] 

#drop duplicate zip codes (there are duplicates due to different property types)
df_zipcodes_submarket_crosswalk  = df_zipcodes_submarket_crosswalk.drop_duplicates(subset='ZIP', keep='first', inplace=False, ignore_index=False) 

#At this point, we have a list of zip codes associated with CoStar markets for our 4 main property types
# print(df_zipcodes_submarket_crosswalk)


# Import Zip Code to County crosswalk (https://www.huduser.gov/portal/datasets/usps_crosswalk.html)
df_zipcodes_county_crosswalk  = pd.read_excel(os.path.join(data_location,'General Data','ZIP_COUNTY_032021.xlsx'), dtype={'ZIP': object,'COUNTY': object} ) 

#Cut down Puerto Rico, Virgin Islands, etc
df_zipcodes_county_crosswalk  = df_zipcodes_county_crosswalk.loc[(df_zipcodes_county_crosswalk['USPS_ZIP_PREF_STATE'] != 'PR') & 
                                                                 (df_zipcodes_county_crosswalk['USPS_ZIP_PREF_STATE'] != 'VI') &
                                                                 (df_zipcodes_county_crosswalk['USPS_ZIP_PREF_STATE'] != 'AS') ]
# print(df_zipcodes_county_crosswalk)


#Merge the submarket crosswalk with the county zip code crosswalk to keep what zip codes are outside of any market
zip_codes_merged = pd.merge(df_zipcodes_county_crosswalk, df_zipcodes_submarket_crosswalk, on='ZIP',how = 'left',indicator='merge_indicator') 
zip_codes_merged = zip_codes_merged.loc[zip_codes_merged['merge_indicator'] == 'left_only'] #keep only the zip codes that don't match with any zip codes in the submarket crosswalk
zip_codes_merged = zip_codes_merged[['ZIP','COUNTY']] #cut down to 2 variables
zip_codes_merged.rename(columns={'COUNTY': 'FIPS Code'}, inplace=True) #Rename county variable
zip_codes_merged  = zip_codes_merged.groupby(['FIPS Code'])['ZIP'].apply(list)
# print(zip_codes_merged)

#At this point, we have a dataframe of all FIPS codes in the US that have at least 1 zip code outside a costar market and a list of the associated "outside market" zips
#we will merge this with a list of US counties to get info on each county such as their name


#Import master list of US Counties
df_master_county_list = pd.read_excel(os.path.join(data_location,'Area Reports Data','County_Master_List.xls'),
                dtype={'FIPS Code': object
                        })

# Merge the dataframe with the zip codes outside any market/submarket with the master county datframe to get the county and state info for that zip code (the zip code dataframe only has the FIPS code)
df_unclassified_zips_with_county_info = pd.merge(zip_codes_merged, df_master_county_list, on='FIPS Code',how = 'left',indicator='merge_indicator')
df_unclassified_zips_with_county_info = df_unclassified_zips_with_county_info.loc[df_unclassified_zips_with_county_info['merge_indicator'] == 'both'] #keep only the zip codes that don't match with any zip codes in the submarket crosswalk
df_unclassified_zips_with_county_info.to_excel(os.path.join(output_location,'Counties Zip Codes Outside CoStar Markets.xlsx'))

print(df_unclassified_zips_with_county_info)