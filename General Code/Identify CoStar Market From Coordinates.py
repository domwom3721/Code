#Date: 03/25/2022
#Author: Mike Leahy
#Summary: Takes a lat lon coordinate (point) as input --> Opens a shapefile with MSA Divisions --> If point falls within a MSA divison ---> 
#         Returns the CBSA code for the division ---> Looks in the CoStar export file for a market with a matching CBSA code ---> 
#         If we find one, reuturn the name of that market  ----> If not, repeat this process with a shapefile of Metropolitan Statistical Areas  --->
#         Either returns a market name or nothing

#CoStar maps their markets to a MSA Divison first if available, then goes to larger MSA if the area has no divisions

#Import packages
from itertools import count
import os 
import shapefile
import pandas as pd
from shapely.geometry import Point
from shapely.geometry.polygon import Polygon

#Specify file paths
dropbox_root                     = os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)')               
msa_divisions_shapefile_location = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'Metro Divisions Shapefile', 'tl_2019_us_metdiv.shp')
msa_shapefile_location           = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'US CBSA Shapefile', 'cb_2018_us_cbsa_500k.shp')
county_shapefile_location        = os.path.join(dropbox_root, 'Research', 'Projects', 'CoStar Map Project', 'US Counties Shapefile', 'cb_2018_us_county_500k.shp')

mf_export                        = os.path.join(dropbox_root, 'Research', 'Projects', 'Research Report Automation Project', 'Data', 'Market Reports Data', 'CoStar Data', 'Raw Data','mf.csv')
office_export                    = os.path.join(dropbox_root, 'Research', 'Projects', 'Research Report Automation Project', 'Data', 'Market Reports Data', 'CoStar Data', 'Raw Data','office.csv')
retail_export                    = os.path.join(dropbox_root, 'Research', 'Projects', 'Research Report Automation Project', 'Data', 'Market Reports Data', 'CoStar Data', 'Raw Data','retail.csv')
industrial_export                = os.path.join(dropbox_root, 'Research', 'Projects', 'Research Report Automation Project', 'Data', 'Market Reports Data', 'CoStar Data', 'Raw Data','industrial.csv')


#Declare the sector and coordinates
sector               = 'Office'
lat, lon             = 40.743864357763115, -74.0310994566193
property_coordinates = Point(lon, lat) #lon, lat

def FindMetroDivCode(point):
    #Open the MSA Division shapefile
    msa_div_map = shapefile.Reader(msa_divisions_shapefile_location)
    

    #Loop through the MSA div map
    for i in range(len(msa_div_map) ):
        polygon               = Polygon(msa_div_map.shape(i).points)
        record                = msa_div_map.shapeRecord(i)
        div_code              = record.record['METDIVFP']
        if  polygon.contains(point):
            return(div_code)

def FindMSACode(point):
    #Open the MSA shapefile
    msa_map     = shapefile.Reader(msa_shapefile_location)
    
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

def FindCoStarMarket(msa_or_div_code, sector):

    #Open the CoStar Export file as a dataframe
    if sector == 'Multifamily':
        costar_df                            = pd.read_csv(mf_export, dtype={'CBSA Code': object})
    elif sector == 'Office':
        costar_df                            = pd.read_csv(office_export, dtype={'CBSA Code': object})
    elif sector == 'Industrial':
        costar_df                            = pd.read_csv(industrial_export, dtype={'CBSA Code': object})
    elif sector == 'Retail':
        costar_df                            = pd.read_csv(retail_export, dtype={'CBSA Code': object})

    #Restrict to our columns of interest
    costar_df = costar_df[['Property Class Name', 'Period', 'Geography Name',	'CBSA Code', 'Geography Type' ]]
    
    #Restrict to Markets (Metros)
    costar_df = costar_df.loc[costar_df['Geography Type'] =='Metro']
    
    #Restrict to Market with matching CBSA code
    costar_df = costar_df.loc[costar_df['CBSA Code'] == msa_or_div_code]

    if len(costar_df) > 0:
        assert len(costar_df) < 50
        market_name = costar_df['Geography Name'].iloc[-1]
        return(market_name)


msa_div_code       = FindMetroDivCode(point=property_coordinates) 
msa_code           = FindMSACode(point=property_coordinates) 
county_fips_code   = FindCountyFips(point=property_coordinates) 

#If we found a msa division code, use that,otherwise use msa code
if msa_div_code != None:
    costar_market = FindCoStarMarket(msa_or_div_code=msa_div_code, sector = sector)

elif msa_div_code == None:
    costar_market = FindCoStarMarket(msa_or_div_code=msa_code, sector = sector)

#Get County name
county_name = CountyNameFromFips(fips=county_fips_code)

print(costar_market)
print(county_name)

















