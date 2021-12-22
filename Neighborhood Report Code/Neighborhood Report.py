#By Mike Leahy
#Started 06/30/2021
#Summary: This script creates reports on neighborhoods/cities for Bowery
import json
# import math
# import re
# from bls_datasets import oes, qcew
# from blsconnect import RequestBLS, bls_search
# from ctypes import addressof
# import census_area
# import docx
# from pprint import pprint
# from numpy import true_divide
# from genericpath import exists
# from us import states
import msvcrt
import os
import re
import shutil
import sys
import time
from datetime import date, datetime

import googlemaps
import mpu
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import pyautogui
import requests
import shapefile
import us
import wikipedia
from bs4 import BeautifulSoup
from census import Census
from census_area import Census as CensusArea
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Row, CT_Tc
from docx.shared import Inches, Pt, RGBColor
from fredapi import Fred
from PIL import Image, ImageOps
from plotly.subplots import make_subplots
from requests.adapters import HTTPAdapter
from requests.exceptions import HTTPError
from requests.packages.urllib3.util.retry import Retry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from shapely.geometry import LineString, MultiPoint, Point, shape,Polygon
from shapely.ops import nearest_points
from walkscore import WalkScoreAPI
from wikipedia.wikipedia import random
from yelpapi import YelpAPI

#Define file paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects', 'Research Report Automation Project') 
main_output_location           =  os.path.join(project_location,'Output','Neighborhood')                   #testing
# main_output_location           =  os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood') #production
data_location                  =  os.path.join(project_location,'Data','Neighborhood Reports Data')
graphics_location              =  os.path.join(project_location,'Data','General Data','Graphics')
map_location                   =  os.path.join(project_location,'Data','Neighborhood Reports Data','Neighborhood Maps')
nyc_cd_map_location            =  os.path.join(project_location,'Data','Neighborhood Reports Data','NYC_CD Maps')


#Data Manipulation functions
def ConvertListElementsToFractionOfTotal(raw_list):
    #Convert list with raw totals into a list where each element is a fraction of the total
    total = sum(raw_list)

    converted_list = []
    for i in raw_list:
        assert i >= 0
        converted_list.append(i/total * 100)
    
    return(converted_list)

def AggregateAcrossDictionaries(neighborhood_tracts_data, fields_list):
    aggregate_dict = {} 
    for field in fields_list:
        total_value = 0

        #Add up all the values from each dictionary
        for d in neighborhood_tracts_data:
            value = d[field]
            total_value = total_value + int(value)
        
        #Add the current field to the new aggregate_dict
        aggregate_dict[field] = total_value
    
    return(aggregate_dict)

def FindMedianCategory(frequency_list, category_list):
    #Takes a list with a fequency distribution (eg [10%,30%,60%]) and the corresponding cateorgories [Red,Blue,Green]
    #Returns the median category, in this case Green
    assert len(frequency_list) == len(category_list)

    total_value_fraction = 0
    for i,value_category_fraction in enumerate(frequency_list):
        total_value_fraction += value_category_fraction
        if total_value_fraction >= 50:
            median_cat_index = i
            break
    
    median_category     = category_list[median_cat_index]
    return(median_category)

#####################################################Geographic Data Related Functions####################################
def GetLatandLon():
    # Look up lat and lon of area with geocoding using google maps api
    gmaps          = googlemaps.Client(key=google_maps_api_key) 
    
    if neighborhood_level == 'custom':
        geocode_result = gmaps.geocode(address=(neighborhood + ', ' + comparison_area + ',' + comparison_state),)
    else:
        geocode_result = gmaps.geocode(address=(neighborhood + ',' + hood_state),)
    
    latitude       = geocode_result[0]['geometry']['location']['lat']
    longitude      = geocode_result[0]['geometry']['location']['lng']

    return([latitude,longitude]) 

def GetNeighborhoodShape():
    if neighborhood_level == 'custom':
        try:
            from osgeo import gdal
            #Method 1: Pull geojson from file with city name
            with open(os.path.join(data_location,'Neighborhood Shapes',comparison_area + '.geojson')) as infile: #Open a geojson file with the city as the name the name of the file with the neighborhood boundries for that city
                my_shape_geojson = json.load(infile)
            
            print('Successfully opened geojson file for ' + comparison_area)
                
            #Iterate through the features in the file (each feature is a negihborhood) and find the boundry of interest
            for i in range(len(my_shape_geojson['features'])):
                feature_hood_name = my_shape_geojson['features'][i]['properties']['name']
                if feature_hood_name == neighborhood:
                    neighborhood_shape = my_shape_geojson['features'][i]['geometry']

            print('Successfully pulled hood shape from stored geojson file')
            return(neighborhood_shape) 
                
        
        except Exception as e:
            print(e,'problem getting shape from ' + comparison_area + ' geojson file')
            print('Looking for exported kml file from my google maps')
            #Method 2: Get bounds from my google maps custom layer export
            
            #Define file locations
            kml_file_download_location         = os.path.join(os.environ['USERPROFILE'],'Downloads', 'Untitled layer.kml')
            kml_file_location                  = os.path.join(data_location,'Neighborhood Shapes',   'Untitled layer.kml')
            new_geojson_file_location          = os.path.join(data_location,'Neighborhood Shapes',   'custom_neighborhood_shape.geojson')
            
            #Step 1: Move the exported kml file from downloads to data folder 
            if os.path.exists(kml_file_download_location) == True:
                print('Moving KML file from downloads folder into data folder')
                shutil.move(kml_file_download_location,kml_file_location)

            #Step 2: Convert the exported google maps kmz file to geojson
            if os.path.exists(new_geojson_file_location) == False:
                print('Converting custom kml file into a geojson file')
                srcDS                              = gdal.OpenEx(kml_file_location)
                ds                                 = gdal.VectorTranslate(new_geojson_file_location, srcDS, format='GeoJSON')

            with open(new_geojson_file_location) as infile: 
                print('Opened geojson file with custom boundraries')
                my_shape_geojson = json.load(infile)
            
            neighborhood_shape       = my_shape_geojson['features'][0]['geometry']
            neighborhood_custom_name = my_shape_geojson['features'][0]['properties']['Name']
            input('We are using a downloaded file from google for custom bounds for ' + neighborhood_custom_name +  ' --- press enter to confirm!')
            return(neighborhood_shape) 

def GetListOfNeighborhoods(city):
    try:
        from osgeo import gdal
            #Method 1: Pull geojson from file with city name
        with open(os.path.join(data_location,'Neighborhood Shapes',city + '.geojson')) as infile: #Open a geojson file with the city as the name the name of the file with the neighborhood boundries for that city
                    my_shape_geojson = json.load(infile)
                
        #Iterate through the features in the file (each feature is a negihborhood) and find the boundry of interest
        feature_hood_names = []
        for i in range(len(my_shape_geojson['features'])):
            feature_hood_name = my_shape_geojson['features'][i]['properties']['name']
            feature_hood_names.append(feature_hood_name)
            
        return(feature_hood_names) 
    except Exception as e:
        print(e)
        return([])

def DetermineNYCCommunityDistrict(lat,lon):
    print('Determining NYC Community District')
    try:
        from osgeo import gdal
        #Method 1: Pull geojson from file with city name
        with open(os.path.join(data_location,'Neighborhood Shapes','NY','nyc_communitydistricts.json')) as infile: #Open a geojson file with the city as the name the name of the file with the neighborhood boundries for that city
            my_shape_geojson = json.load(infile)
        
        #Iterate through the features in the file (each feature is a communtiy district) and find the boundry of interest
        for communtiy_district in range(len(my_shape_geojson['features'])):
            communtiy_district_number = my_shape_geojson['features'][communtiy_district]['properties']["BoroCD"]
            communtiy_district_shape  = my_shape_geojson['features'][communtiy_district]['geometry']['coordinates'][0]
            
            try:
                point   = Point(lon, lat)
                polygon = Polygon([tuple(l) for l in communtiy_district_shape])      
                                
                #Check if lat and lon is inside the communtiy district
                if polygon.contains(point) == True:
                    return(communtiy_district_number)

            except Exception as e:
                print(e)
                continue
        
        return('x')

    except Exception as e:
        print(e)
        return('x')

#####################################################User FIPS input proccessing Functions####################################

def ProcessPlaceFIPS(place_fips):
    #This function takes a user provided 7 digit census place fips code and returns a list of key variables about that fips code
    #eg: the place name, type, state name, state code, etc

    #Process the FIPS code provided    
    place_fips                      = place_fips.replace('-','').strip()
    assert len(place_fips)          == 7
    state_fips                      = place_fips[0:2]
    place_fips                      = place_fips[2:]

    #Get name of the hood using the FIPS code provided
    place_name                      = c.sf1.state_place(fields=['NAME'], state_fips = state_fips, place = place_fips)[0]['NAME']
    state_full_name                 = place_name.split(',')[1].strip()
    place_name                      = place_name.split(',')[0].strip()
    place_type                      = place_name.split(' ')[len(place_name.split(' '))-1] #eg: village, city, etc
    place_name                      = ' '.join(place_name.split(' ')[0:len(place_name.split(' '))-1]).title()
    
    #Name of State
    state                           = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
    state                           = state.abbr
    assert len(state)               == 2

    return([place_fips, state_fips, place_name, state_full_name, state,place_type])

def ProcessCountyFIPS(county_fips):
    
    #Process the FIPS code provided by user
    county_fips               = county_fips.replace('-','').strip()
    assert len(county_fips) == 5
    state_fips               = county_fips[0:2]
    county_fips              = county_fips[2:]

    #Get name of county
    name                     = c.sf1.state_county(fields=['NAME'], state_fips = state_fips, county_fips = county_fips)[0]['NAME']
    state_full_name          = name.split(',')[1].strip()
    name                     = name.split(',')[0].strip().title()

    #Name of State
    state                   = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
    state                   = state.abbr

    assert len(state)       == 2

    return[county_fips, state_fips, name, state_full_name, state]

def ProcessCountySubdivisionFIPS(county_subdivision_fips):
    #Proccess FIPS code provided
    county_subdivision_fips = county_subdivision_fips.replace('-','').strip()
    assert len(county_subdivision_fips) == 10
    state_fips       = county_subdivision_fips[0:2]
    county_fips      = county_subdivision_fips[2:5]
    suvdiv_fips      = county_subdivision_fips[5:]

    #Get name of hood using the FIPS code provided
    name             = c.sf1.state_county_subdivision(fields=['NAME'],state_fips = state_fips,county_fips = county_fips, subdiv_fips = suvdiv_fips)[0]['NAME']
    state_full_name  = name.split(',')[2].strip()
    name             = name.split(',')[0].strip().title()
    place_type       = name.split(' ')[len(name.split(' '))-1] #eg: village, city, etc
    name             = ' '.join(name.split(' ')[0:len(name.split(' '))-1]).title()

    #Name of State
    state            = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
    state            = state.abbr
    assert len(state) == 2

    return([suvdiv_fips,county_fips,name,state_fips,state_full_name,state,place_type])

def ProcessCountyTract(tract,county_fips):
    #Takes a user provided county fips code and a census tract number and returns a list of key variables
    county_fips               = county_fips.replace('-','').strip()
    assert len(county_fips)   == 5
    state_fips                = county_fips[0:2]
    county_fips               = county_fips[2:]

    tract                     = tract.replace('-','').strip()
    assert len(tract)         == 6

    #Get name of tract
    name                      = c.sf1.state_county_tract(fields=['NAME'],state_fips = state_fips, county_fips = county_fips,tract = tract)[0]['NAME']
    state_full_name           = name.split(',')[2].strip()
    name                      = name.split(',')[0] + ',' +  name.split(',')[1]
    name                      = name.strip().title()

    #Name of State
    state                     = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
    state                     = state.abbr
    assert len(state)         == 2

    return([county_fips, tract, name, state_full_name, state, state_fips])

def ProcessZipCode(zip_code):
    #Process the zip code provided
    zip_code                            = str(zip_code).replace('-','').strip()
    assert len(zip_code) == 5
    
    #Get the state FIPS code (eg New York: 36)
    zip_county_crosswalk_df            = pd.read_excel(os.path.join(data_location,'Census Area Codes','ZIP_COUNTY_092021.xlsx')) #read in crosswalk file
    zip_county_crosswalk_df['ZIP']     = zip_county_crosswalk_df['ZIP'].astype(str)
    zip_county_crosswalk_df['ZIP']     = zip_county_crosswalk_df['ZIP'].str.zfill(5)
    zip_county_crosswalk_df['COUNTY']  = zip_county_crosswalk_df['COUNTY'].astype(str)
    zip_county_crosswalk_df['COUNTY']  = zip_county_crosswalk_df['COUNTY'].str.zfill(5)

    zip_county_crosswalk_df            = zip_county_crosswalk_df.loc[zip_county_crosswalk_df['ZIP'] == zip_code]                 #restrict to rows for zip code
    county_fips                        = str(zip_county_crosswalk_df['COUNTY'].iloc[-1])[2:]
    state_fips                         = str(zip_county_crosswalk_df['COUNTY'].iloc[-1])[0:2] #Get state fips from the county fips code (the county the zip code is in)
    assert len(state_fips) == 2

    #Get name of hood
    name                               = c.sf1.state_zipcode(fields=['NAME'],state_fips=state_fips, zcta=zip_code)[0]['NAME']
    state_full_name                    = name.split(',')[1].strip()
    name                               = name.split(',')[0].replace('ZCTA5','').strip().title() + ' (Zip Code)'

    #Name of State
    state                              = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
    state                              = state.abbr
    assert                 len(state) == 2


    return([county_fips, zip_code, name,state_full_name,state,state_fips])

def PlaceFIPSToCountyFIPS(place_fips):
    #Takes 7 digit place fips code for a city and returns the 5 digit fips code for that city
    
    #Open file with place fips code and county fips code
    place_county_crosswalk_df                            = pd.read_csv(os.path.join(data_location,'Census Area Codes','national_places.csv'),encoding='latin-1') #read in crosswalk file
    
    place_county_crosswalk_df['PLACEFP']                 = place_county_crosswalk_df['PLACEFP'].astype(str)
    place_county_crosswalk_df['PLACEFP']                 = place_county_crosswalk_df['PLACEFP'].str.zfill(5)
    place_county_crosswalk_df['County_FIPS']             = place_county_crosswalk_df['County_FIPS'].astype(str)
    place_county_crosswalk_df['County_FIPS']             = place_county_crosswalk_df['County_FIPS'].str.zfill(7)


    #Restrict to observations that include the provieded place fips
    place_county_crosswalk_df            = place_county_crosswalk_df.loc[place_county_crosswalk_df['PLACEFP'] == str(place_fips)]                 #restrict to rows for zip code
    
    #Return the last row if that's there's only one, otherwise ask user to choose
    if len(place_county_crosswalk_df) == 1:
        county_fips                         = str(place_county_crosswalk_df['County_FIPS'].iloc[-1])[0:5]
    elif len(place_county_crosswalk_df) > 1:
        print(place_county_crosswalk_df)
        input('There are more than 1 counties for this city: using last one please confirm') #** Fix later by giving user chance to choose 
        county_fips                         = str(place_county_crosswalk_df['County_FIPS'].iloc[-1])[0:5]
    else:
        return(None)


    return(county_fips)
         
#####################################################Misc Functions####################################
def CreateDirectory():
    print('Creating Directories and file name')
    global report_path,hood_folder_map,hood_folder
    
    state_folder_map         = os.path.join(map_location,hood_state)

    state_folder             = os.path.join(main_output_location,hood_state)

    if neighborhood_level == 'custom':

        if os.path.exists(state_folder) == False:
            os.mkdir(state_folder)  
        
        if os.path.exists(state_folder_map) == False:
            os.mkdir(state_folder_map)  

        city_folder =  os.path.join(main_output_location,hood_state,comparison_area)
        city_folder_map =  os.path.join(map_location,hood_state,comparison_area)

        if os.path.exists(city_folder) == False:
            os.mkdir(city_folder) 
        
        if os.path.exists(city_folder_map) == False:
            os.mkdir(city_folder_map) 


        hood_folder              = os.path.join(main_output_location,hood_state,comparison_area,neighborhood)
        hood_folder_map          = os.path.join(map_location,hood_state,city_folder_map,neighborhood)


    else:
        hood_folder              = os.path.join(main_output_location,hood_state,neighborhood)
        hood_folder_map          = os.path.join(map_location,hood_state,neighborhood)
    



    for folder in [state_folder,hood_folder,state_folder_map,hood_folder_map]:
         if os.path.exists(folder):
            pass 
         else:
            os.mkdir(folder) 
    
    report_path = os.path.join(hood_folder,current_year + ' ' + hood_state + ' - ' + neighborhood  + ' - hood' + '_draft')[:255] 
    report_path = report_path + '.docx'

def FindZipCodeDictionary(zip_code_data_dictionary_list,zcta,state_fips):
    #This function takes a list of dictionaries, where each zip code gets its own dictionary. Takes a zip code and state fips code and finds and returns just that dictionary.
    #We need to use this, because the census api is causing an error that requires us to retrive data for all zip codes in the country
    for zcta_dictionary in  zip_code_data_dictionary_list:
    
        if zcta_dictionary['zip code tabulation area'] == zcta and zcta_dictionary['state'] == state_fips:
            return(zcta_dictionary)
        

    print('Could not find dictionary for given zip code: ', zcta )

#Data Gathering Related Functions
def DeclareAPIKeys():
    global census_api_key,walkscore_api_key,google_maps_api_key,yelp_api_key,yelp_api,yelp_client_id,location_iq_api_key
    global c,c_area,walkscore_api
    global zoneomics_api_key
    
    #Declare API Keys
    census_api_key                = '18335344cf4a0242ae9f7354489ef2f8860a9f61'
    walkscore_api_key             = '057f7c0a590efb7ec06da5a8735e536d'
    google_maps_api_key           = 'AIzaSyBMcoRFOW2rxAGxURCpA4gk10MROVVflLs'
    yelp_client_id                = 'NY9c0_9kvOU4wfzmkkruOQ'
    yelp_api_key                  = 'l1WjEgdgSMpU9PJtXEk0bLs4FJdsVLONqJLhbaA0gZlbFyEFUTTkxgRzBDc-_5234oLw1CLx-iWjr8w4nK_tZ_79qVIOv3yEMQ9aGcSS8xO1gkbfENCBKEl34COVYXYx'
    location_iq_api_key           = 'pk.8937271b8b15004065ca62552e7d06f7'
    zoneomics_api_key             = 'd69b3eee92f8d3cec8c71893b340faa8cb52e1b8'

    yelp_api                      = YelpAPI(yelp_api_key)
    walkscore_api                 = WalkScoreAPI(api_key = walkscore_api_key)
    c                             = Census(census_api_key) #Census API wrapper package
    c_area                        = CensusArea(census_api_key) #Census API package, sepearete extension of main package that allows for custom boundries

#Data Gathering Related Functions
def DeclareFormattingParameters():
    global primary_font
    global primary_space_after_paragraph
    
    #Set formatting paramaters for reports
    primary_font                  = 'Avenir Next LT Pro Light' 
    primary_space_after_paragraph = 8
class TimeoutExpired(Exception):
    pass

def input_with_timeout(prompt, timeout, timer=time.monotonic):
    sys.stdout.write(prompt)
    sys.stdout.flush()
    endtime = timer() + timeout
    result = []
    while timer() < endtime:
        if msvcrt.kbhit():
            result.append(msvcrt.getwche()) #XXX can it block on multibyte characters?
            if result[-1] == '\r':
                return ''.join(result[:-1])
        time.sleep(0.04) # just to yield to other processes/threads
    raise TimeoutExpired

#####################################################Census Data Related Functions####################################
def GetCensusFrequencyDistribution(geographic_level,hood_or_comparison_area,fields_list,operator):
    #A general function that takes a list of census variables that represent a set of all possible categoreis (eg: a list of home value categories)
    #It then creates a list with the number of observations in each cateogry,
    #It then converts the total ammount elements to fractions of the total
    
    #The basic mechanics are this ['men','women'] ----> [30,70] ----> [.30,.70]

    #Speicify geographic level specific varaibles
    if geographic_level == 'place':
        try:

            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips
                state_fips = hood_state_fips
            
            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparison_place_fips
                state_fips = comparison_state_fips
            
            neighborhood_household_size_distribution_raw = operator.state_place(fields = fields_list,state_fips = state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area)
            return()
    
    elif geographic_level == 'county':
        
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                state_fips  = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                state_fips  = comparison_state_fips

            neighborhood_household_size_distribution_raw =operator.state_county(fields = fields_list, state_fips = state_fips,county_fips = county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips
                state_fips  = hood_state_fips


            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
                state_fips = comparison_state_fips

            neighborhood_household_size_distribution_raw = operator.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        except Exception as e:
            print(e, 'Problem getting data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip
                state_fips  = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                zcta       = comparison_zip
                state_fips = comparison_state_fips

            neighborhood_household_size_distribution_raw = operator.state_zipcode(fields=fields_list,state_fips=state_fips,zcta=zcta)[0]
        except Exception as e:
            print(e, 'Problem getting data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips
                state_fips  = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
                state_fips  = comparison_state_fips
            
            neighborhood_household_size_distribution_raw = operator.state_county_tract(fields=fields_list, state_fips = state_fips,county_fips=county_fips,tract=tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        if operator == c.acs5:
            operator = c_area.acs5
        elif operator == c.sf1:
            operator = c_area.sf1

        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = operator.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        neighborhood_household_size_distribution_raw = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )
    
    
    #General data manipulation (same for all geographic levels)
    distribution = []
    for field in fields_list:
            distribution.append(neighborhood_household_size_distribution_raw[field])

    try:    
        distribution = ConvertListElementsToFractionOfTotal(distribution)
        return(distribution)
    except Exception as e:
        print(e)

#Households by number of memebrs
def GetHouseholdSizeData(geographic_level,hood_or_comparison_area):
    print('Getting household size data for: ',hood_or_comparison_area)
    neighborhood_household_size_distribution       = GetCensusFrequencyDistribution(     geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area,fields_list=['H013002','H013003','H013004','H013005','H013006','H013007','H013008'],operator=c.sf1)          #Neighborhood households by size
    return(neighborhood_household_size_distribution)
    
#Household Tenure (owner-occupied vs renter-occupied)
def GetHousingTenureData(geographic_level,hood_or_comparison_area):
    #Occupied Housing Units by Tenure
    print('Getting tenure data for: ',hood_or_comparison_area)
    neighborhood_tenure_distribution  = GetCensusFrequencyDistribution(     geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area,fields_list = ['H004004','H004003','H004002'],operator=c.sf1) 

    #add together the owned free and clear percentage with the owned with a mortgage percentage to simply an owner-occupied fraction
    neighborhood_tenure_distribution[1] = neighborhood_tenure_distribution[1] +  neighborhood_tenure_distribution[2]
    del neighborhood_tenure_distribution[2]
    return(neighborhood_tenure_distribution)

#Housing related data functions
def GetHousingValues(geographic_level,hood_or_comparison_area):
    print('Getting housing value data for: ',hood_or_comparison_area)
    household_value_data = GetCensusFrequencyDistribution(geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area, fields_list = ["B25075_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,28)],operator=c.acs5)  
    return(household_value_data)

#Year Housing Built Data
def GetHouseYearBuiltData(geographic_level,hood_or_comparison_area):
    print('Getting year built data for: ',hood_or_comparison_area)
    year_built_data = GetCensusFrequencyDistribution(    geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area,  fields_list = ["B25034_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,12)],operator= c_area.acs5)
    year_built_data.reverse()
    return(year_built_data)

#Travel Method to work
def GetTravelMethodData(geographic_level,hood_or_comparison_area):
    print('Getting travel method to work data for: ' + hood_or_comparison_area)
    neighborhood_method_to_work_distribution = GetCensusFrequencyDistribution(geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area, fields_list =['B08006_003E','B08006_004E','B08006_008E','B08006_015E','B08006_017E','B08006_014E','B08006_016E'],operator=c.acs5)  
    return(neighborhood_method_to_work_distribution) 

#Household Income data functions
def GetHouseholdIncomeValues(geographic_level,hood_or_comparison_area):
    print('Getting household income data for: ',hood_or_comparison_area)
    total_income_breakdown = GetCensusFrequencyDistribution( geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area, fields_list = ["B19001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,18)],operator=c.acs5) 
    return(total_income_breakdown)

#Travel Time to Work
def GetTravelTimeData(geographic_level,hood_or_comparison_area):
    print('Getting travel time data for: ', hood_or_comparison_area)
     #5 Year ACS travel time range:   B08012_003E - B08012_013E
    travel_time_data = GetCensusFrequencyDistribution(        geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area,fields_list = ["B08012_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,14)],operator=c.acs5)   
    return(travel_time_data)

#Age Related Data Functions
def GetAgeData(geographic_level,hood_or_comparison_area):
    print('Getting age breakdown for: ',hood_or_comparison_area)
    #Return a list with the fraction of the population in different age groups 

    #Define 2 lists of variables, 1 for male age groups and another for female
    male_fields_list   =  ["B01001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(3,26)]  #5 Year ACS age variables for men range:  B01001_003E - B01001_025E
    female_fields_list =  ["B01001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(27,50)] #5 Year ACS age variables for women range:  B01001_027E - B01001_049E

   
    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips
                state_fips = hood_state_fips


            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparison_place_fips
                state_fips = comparison_state_fips

            male_age_data = c.acs5.state_place(fields=male_fields_list, state_fips=state_fips,place=place_fips)[0]
            female_age_data = c.acs5.state_place(fields=female_fields_list,state_fips=state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                state_fips  = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                state_fips  = comparison_state_fips

            male_age_data   = c.acs5.state_county(fields=male_fields_list,state_fips=state_fips,county_fips=county_fips)[0]
            female_age_data = c.acs5.state_county(fields=female_fields_list,state_fips=state_fips,county_fips=county_fips)[0]

        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips
                state_fips  = hood_state_fips


            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
                state_fips  = comparison_state_fips


            male_age_data   = c.acs5.state_county_subdivision(fields=male_fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
            female_age_data = c.acs5.state_county_subdivision(fields=female_fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
            try:        
                if hood_or_comparison_area == 'hood':
                    zcta = hood_zip

                elif hood_or_comparison_area == 'comparison area':
                    zcta = comparison_zip
            
                male_age_data       = c.acs5.zipcode(fields = male_fields_list, zcta = '*')
                male_age_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   male_age_data  , zcta = zcta, state_fips = state_fips )

                female_age_data       = c.acs5.zipcode(fields = female_fields_list, zcta = '*')
                female_age_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   female_age_data  , zcta = zcta, state_fips = state_fips )

            except Exception as e:
                print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
                return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips


            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips

            male_age_data = c.acs5.state_county_tract(fields=male_fields_list,state_fips=state_fips,county_fips=county_fips, tract=tract)[0]
            female_age_data = c.acs5.state_county_tract(fields=female_fields_list,state_fips=state_fips,county_fips=county_fips, tract=tract)[0]
        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
        
    elif geographic_level == 'custom':
        
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_male_tracts_data   = []
        neighborhood_female_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_male_census_data   = c_area.acs5.geo_tract(male_fields_list, neighborhood_shape)
        raw_female_census_data = c_area.acs5.geo_tract(female_fields_list, neighborhood_shape)
        

        for tract_geojson, tract_data, tract_proportion in raw_male_census_data:
            neighborhood_male_tracts_data.append((tract_data))
        
        for tract_geojson, tract_data, tract_proportion in raw_female_census_data:
            neighborhood_female_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        male_age_data   = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_male_tracts_data, fields_list   = male_fields_list )
        female_age_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_female_tracts_data, fields_list = female_fields_list )

    #Create an empty list and place the age values from the dictionary inside of it
    male_age_breakdown = []
    for field in male_fields_list:
        male_age_breakdown.append(male_age_data[field])


    #Create an empty list and place the age values from the dictionary inside of it
    female_age_breakdown = []
    for field in female_fields_list:
        female_age_breakdown.append(female_age_data[field])
    

    total_age_breakdown = []
    for (men, women) in zip(male_age_breakdown, female_age_breakdown):
        total = (men + women)
        total_age_breakdown.append(total)

    
    #Consolidate some of the age groups into larger groups
    total_age_breakdown[0] = sum(total_age_breakdown[0:5])
    total_age_breakdown[1] = sum(total_age_breakdown[5:8])
    total_age_breakdown[2] = sum(total_age_breakdown[8:10])
    total_age_breakdown[3] = sum(total_age_breakdown[10:13])
    total_age_breakdown[4] = sum(total_age_breakdown[13:18])
    total_age_breakdown[5] = sum(total_age_breakdown[18:])
    del[total_age_breakdown[6:]]


    #Convert from raw numbers to fractions of total
    total_age_breakdown = ConvertListElementsToFractionOfTotal(total_age_breakdown)

    return(total_age_breakdown)

#Number of Housing Units based on number of units in building
def GetNumberUnitsData(geographic_level,hood_or_comparison_area):
    print('Getting housing units by number of units data for: ', hood_or_comparison_area)
    
    
    owner_occupied_fields_list  = ["B25032_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(3,11)]   #5 Year ACS owner occupied number of units variables range:  B25032_003E - B25032_010E
    renter_occupied_fields_list = ["B25032_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(14,22)]  #5 Year ACS renter occupied number of units variables range: B25032_014E - B25032_021E 

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips
                state_fips = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparison_place_fips
                state_fips = comparison_state_fips
        
            owner_occupied_units_raw_data = c.acs5.state_place(fields = owner_occupied_fields_list,state_fips=state_fips,place=place_fips)[0]
            renter_occupied_units_raw_data = c.acs5.state_place(fields = renter_occupied_fields_list,state_fips=state_fips,place=place_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county':
        try:
            
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                state_fips = hood_state_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                state_fips = comparison_state_fips

            owner_occupied_units_raw_data  = c.acs5.state_county(fields = owner_occupied_fields_list,  state_fips = state_fips, county_fips = county_fips)[0]
            renter_occupied_units_raw_data = c.acs5.state_county(fields = renter_occupied_fields_list, state_fips = state_fips, county_fips = county_fips)[0]
        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips
                state_fips  = hood_state_fips


            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
                state_fips = comparison_state_fips

        
            owner_occupied_units_raw_data  = c.acs5.state_county_subdivision(fields = owner_occupied_fields_list, state_fips  = state_fips, county_fips=county_fips,  subdiv_fips=subdiv_fips)[0]
            renter_occupied_units_raw_data = c.acs5.state_county_subdivision(fields = renter_occupied_fields_list, state_fips = state_fips, county_fips=county_fips,  subdiv_fips=subdiv_fips)[0]
        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:        
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
        
            owner_occupied_units_raw_data       = c.acs5.zipcode(fields = owner_occupied_fields_list,  zcta = '*' )
            owner_occupied_units_raw_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   owner_occupied_units_raw_data  , zcta = zcta, state_fips = state_fips )

            renter_occupied_units_raw_data      = c.acs5.zipcode(fields = renter_occupied_fields_list, zcta = '*' )
            renter_occupied_units_raw_data      = FindZipCodeDictionary(zip_code_data_dictionary_list =   renter_occupied_units_raw_data  , zcta = zcta, state_fips = state_fips )


        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips

            owner_occupied_units_raw_data  = c.acs5.state_county_tract(fields = owner_occupied_fields_list, state_fips=state_fips, county_fips=county_fips,  tract=tract)[0]
            renter_occupied_units_raw_data = c.acs5.state_county_tract(fields = renter_occupied_fields_list, state_fips=state_fips, county_fips=county_fips, tract=tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_oo_tracts_data   = []
        neighborhood_ro_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_oo_census_data   = c_area.acs5.geo_tract(owner_occupied_fields_list, neighborhood_shape)
        raw_ro_census_data = c_area.acs5.geo_tract(renter_occupied_fields_list, neighborhood_shape)
        

        for tract_geojson, tract_data, tract_proportion in raw_oo_census_data:
            neighborhood_oo_tracts_data.append((tract_data))
        
        for tract_geojson, tract_data, tract_proportion in raw_ro_census_data:
            neighborhood_ro_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        owner_occupied_units_raw_data   = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_oo_tracts_data, fields_list   = owner_occupied_fields_list )
        renter_occupied_units_raw_data  = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_ro_tracts_data, fields_list = renter_occupied_fields_list )

    
    


    #Create an empty list and place the values from the dictionary inside of it
    owner_occupied_units_data = []
    for field in owner_occupied_fields_list:
        owner_occupied_units_data.append(owner_occupied_units_raw_data[field])

    #Now repeat for the renter occupied fields
    #Create an empty list and place the values from the dictionary inside of it
    renter_occupied_units_data = []
    for field in renter_occupied_fields_list:
        renter_occupied_units_data.append(renter_occupied_units_raw_data[field])

    
    total_units = sum(owner_occupied_units_data) + sum(renter_occupied_units_data)

    total_unit_data = []
    for (oo, ro) in zip(owner_occupied_units_data, renter_occupied_units_data):
        total_unit_data.append(( (oo + ro )/total_units) * 100)

    
    return(total_unit_data)

#Occupations Data
def GetTopOccupationsData(geographic_level,hood_or_comparison_area):
    print('Getting occupation data for: ',hood_or_comparison_area)
    top_occupations_data = GetCensusFrequencyDistribution(        geographic_level = geographic_level, hood_or_comparison_area = hood_or_comparison_area,fields_list =  ['B24011_002E','B24011_018E','B24011_026E','B24011_029E','B24011_036E'],operator=c.acs5)   
    return(top_occupations_data)
    
def GetOverviewTable(hood_geographic_level,comparison_geographic_level):
    print('Getting Overview table data')

    total_pop_field               = 'P001001'
    total_households_field        = 'H003002'

    acs_total_pop_field           = 'B01001_001E'
    acs_total_households_field    = 'B09005_001E'  

    redistricting_total_pop_field = 'P1_001N'
    redistricting_total_hh_field  = 'H1_002N'

    #calculate table variables for hood
    if hood_geographic_level == 'place':
        current_estimate_period = '2020 Census'

        _2010_hood_pop         = c.sf1.state_place(fields = total_pop_field,                                        state_fips = hood_state_fips, place = hood_place_fips)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_place(fields = total_households_field,                                 state_fips = hood_state_fips, place = hood_place_fips)[0][total_households_field]
        
        current_hood_pop       = c.pl.state_place(fields = [redistricting_total_pop_field],                          state_fips = hood_state_fips, place = hood_place_fips)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_place(fields = [redistricting_total_hh_field],                           state_fips = hood_state_fips, place = hood_place_fips)[0][redistricting_total_hh_field]
        
    elif hood_geographic_level == 'county':
        current_estimate_period = '2020 Census'

        _2010_hood_pop   = c.sf1.state_county(fields = total_pop_field,                      state_fips = hood_state_fips, county_fips = hood_county_fips)[0][total_pop_field]
        _2010_hood_hh    = c.sf1.state_county(fields = total_households_field,               state_fips = hood_state_fips, county_fips = hood_county_fips)[0][total_households_field]
    
        current_hood_pop =  c.pl.state_county(fields = redistricting_total_pop_field,        state_fips = hood_state_fips, county_fips = hood_county_fips)[0][redistricting_total_pop_field]
        current_hood_hh  =  c.pl.state_county(fields = redistricting_total_hh_field,         state_fips = hood_state_fips, county_fips = hood_county_fips)[0][redistricting_total_hh_field]
        
    elif hood_geographic_level == 'county subdivision':
        current_estimate_period = '2020 Census'
        _2010_hood_pop         = c.sf1.state_county_subdivision(fields = total_pop_field,                     state_fips = hood_state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_county_subdivision(fields = total_households_field,              state_fips = hood_state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][total_households_field]

        current_hood_pop       = c.pl.state_county_subdivision(fields = redistricting_total_pop_field,        state_fips = hood_state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_county_subdivision(fields = redistricting_total_hh_field,         state_fips = hood_state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][redistricting_total_hh_field]

    elif hood_geographic_level == 'zip':
        current_estimate_period = 'Current Estimate'

        _2010_hood_pop         = c.sf1.state_zipcode(fields = total_pop_field,        state_fips = hood_state_fips, zcta = hood_zip)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_zipcode(fields = total_households_field, state_fips = hood_state_fips, zcta = hood_zip)[0][total_households_field]

        current_hood_pop       = _2010_hood_pop
        current_hood_hh        = _2010_hood_hh

    elif hood_geographic_level == 'tract':
        current_estimate_period = '2020 Census'
        _2010_hood_pop         = c.sf1.state_county_tract(fields = total_pop_field,              state_fips = hood_state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_county_tract(fields = total_households_field,       state_fips = hood_state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][total_households_field]

        current_hood_pop       = c.pl.state_county_tract(fields = redistricting_total_pop_field, state_fips = hood_state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_county_tract(fields = redistricting_total_hh_field,  state_fips = hood_state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][redistricting_total_hh_field]

    elif hood_geographic_level == 'custom':
        current_estimate_period = 'Current Estimate'
        
        #2010 Population
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.sf1.geo_tract(total_pop_field, neighborhood_shape)
       
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        _2010_hood_pop_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = [total_pop_field])
        _2010_hood_pop          = _2010_hood_pop_raw_data[total_pop_field]


        #2010 Households
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.sf1.geo_tract(total_households_field, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        _2010_hood_hh_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = [total_households_field])
        _2010_hood_hh          = _2010_hood_hh_raw_data[total_households_field]


        # current_hood_pop       = _2010_hood_pop
        # current_hood_hh        = _2010_hood_hh



        #2019 Population
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(acs_total_pop_field, neighborhood_shape)
       
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        current_hood_pop_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = [acs_total_pop_field])
        current_hood_pop          = current_hood_pop_raw_data[acs_total_pop_field]


        #2019 HH
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(acs_total_households_field, neighborhood_shape)
       
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        current_hood_hh_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = [acs_total_households_field])
        current_hood_hh          = current_hood_hh_raw_data[acs_total_households_field]













        # #2020 Population and Households

        # #Start by getting all of the tracts within the custom area
        # neighborhood_tracts_list          = [] #list for each tract's number
        # neighborhood_tracts_counties_list = [] #list for each tract's county fips


        # raw_tracts_list = c_area.sf1.geo_tract('NAME', neighborhood_shape)
       
        # for tract_geojson, tract_data, tract_proportion in raw_tracts_list:
        #     neighborhood_tracts_list.append((tract_data['tract']))
        #     neighborhood_tracts_counties_list.append((tract_data['county']))

            
        # #Now loop through list of tracts and get 20220 redistricting data population for each tract, add that value to the totals
        # current_hood_pop       = 0
        # current_hood_hh        = 0

        # for tract,hood_county_fips in zip(neighborhood_tracts_list, neighborhood_tracts_counties_list):
        #     print(tract)
        #     tract_pop        = c.pl.state_county_tract(fields = 'NAME',        state_fips = state_fips, county_fips = hood_county_fips, tract = tract) #[0][redistricting_total_pop_field]
        #     tract_hh         = c.pl.state_county_tract(fields = 'NAME',         state_fips = state_fips, county_fips = hood_county_fips, tract = tract) #[0][redistricting_total_hh_field]
            
        #     print(tract_pop)
        #     print(tract_hh)

        #     current_hood_pop += tract_pop
        #     current_hood_hh  += tract_hh

    #Table variables for comparison area
    if comparison_geographic_level == 'place':
        _2010_comparison_pop = c.sf1.state_place(fields = total_pop_field,                       state_fips = comparison_state_fips, place = comparison_place_fips)[0][total_pop_field]
        _2010_comparison_hh  = c.sf1.state_place(fields = total_households_field,                state_fips = comparison_state_fips, place = comparison_place_fips)[0][total_households_field]

        current_comparison_pop = c.pl.state_place(fields = redistricting_total_pop_field,        state_fips = comparison_state_fips, place = comparison_place_fips)[0][redistricting_total_pop_field]
        current_comparison_hh  = c.pl.state_place(fields = redistricting_total_hh_field,         state_fips = comparison_state_fips, place = comparison_place_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'county':
        _2010_comparison_pop   = c.sf1.state_county(fields = total_pop_field,                      state_fips = comparison_state_fips, county_fips = comparison_county_fips)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_county(fields = total_households_field,               state_fips = comparison_state_fips, county_fips = comparison_county_fips)[0][total_households_field]

        current_comparison_pop =  c.pl.state_county(fields = redistricting_total_pop_field,        state_fips = comparison_state_fips, county_fips = comparison_county_fips)[0][redistricting_total_pop_field]
        current_comparison_hh  =  c.pl.state_county(fields = redistricting_total_hh_field,         state_fips = comparison_state_fips, county_fips = comparison_county_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'county subdivision':
        _2010_comparison_pop    = c.sf1.state_county_subdivision(fields = total_pop_field,                     state_fips = comparison_state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][total_pop_field]
        _2010_comparison_hh     = c.sf1.state_county_subdivision(fields = total_households_field,              state_fips = comparison_state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][total_households_field]

        current_comparison_pop  = c.pl.state_county_subdivision(fields = redistricting_total_pop_field,        state_fips = comparison_state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][redistricting_total_pop_field]
        current_comparison_hh   = c.pl.state_county_subdivision(fields = redistricting_total_hh_field,         state_fips = comparison_state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'zip':
        _2010_comparison_pop   = c.sf1.state_zipcode(fields = total_pop_field,state_fips=comparison_state_fips,zcta = comparison_zip)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_zipcode(fields = total_households_field,state_fips=comparison_state_fips,zcta=comparison_zip)[0][total_households_field]

        current_comparison_pop = _2010_comparison_pop
        current_comparison_hh  = _2010_comparison_hh

    elif comparison_geographic_level == 'tract':
        _2010_comparison_pop   = c.sf1.state_county_tract(fields = total_pop_field,                     state_fips = comparison_state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_county_tract(fields = total_households_field,              state_fips = comparison_state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][total_households_field]

        current_comparison_pop =  c.pl.state_county_tract(fields = redistricting_total_pop_field,        state_fips = comparison_state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][redistricting_total_pop_field]
        current_comparison_hh  =  c.pl.state_county_tract(fields = redistricting_total_hh_field,         state_fips = comparison_state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'custom':
        pass



    
     #Stand in for current pop estiamtes
    


    #Calculate growth rates
    hood_pop_growth        = ((int(current_hood_pop)/int(_2010_hood_pop)) - 1) * 100
    hood_hh_growth         = ((int(current_hood_hh)/int(_2010_hood_hh))   - 1) * 100
    comparsion_pop_growth  =  (int(current_comparison_pop)/int(_2010_comparison_pop) - 1) * 100
    comparsion_hh_growth   =  (int(current_comparison_hh)/int(_2010_comparison_hh)   - 1) * 100

    #Format pop and hh variables
    _2010_hood_pop         = "{:,.0f}".format(int(_2010_hood_pop))
    _2010_hood_hh          = "{:,.0f}".format(int(_2010_hood_hh))
    _2010_comparison_pop   = "{:,.0f}".format(int(_2010_comparison_pop))
    _2010_comparison_hh    = "{:,.0f}".format(int(_2010_comparison_hh))
    current_hood_pop       = "{:,.0f}".format(int(current_hood_pop))
    current_hood_hh        = "{:,.0f}".format(int(current_hood_hh))
    current_comparison_pop = "{:,.0f}".format(int(current_comparison_pop))
    current_comparison_hh  = "{:,.0f}".format(int(current_comparison_hh))

    #Format growth variables
    hood_pop_growth         = "{:,.1f}%".format(hood_pop_growth)
    hood_hh_growth          = "{:,.1f}%".format(hood_hh_growth)
    comparsion_pop_growth   = "{:,.1f}%".format(comparsion_pop_growth)
    comparsion_hh_growth    = "{:,.1f}%".format(comparsion_hh_growth)


    #each row represents a row of data for overview table
    row1 = [''          , 'Area',             '2010 Census',            current_estimate_period,                                      '% Change']
    row2 = ['Population', neighborhood,        _2010_hood_pop,          current_hood_pop ,                                 hood_pop_growth ]
    row3 = [''          , comparison_area,     _2010_comparison_pop,    current_comparison_pop,                       comparsion_pop_growth]
    row4 = ['Households', neighborhood,        _2010_hood_hh,           current_hood_hh,                                     hood_hh_growth]
    row5 = [''          , comparison_area,     _2010_comparison_hh,     current_comparison_hh,                        comparsion_hh_growth ]
    
    return(    
            [ 
             row1,
             row2,
             row3,
             row4,
             row5
                  ]
         )

#####################################################Non Census Sources Data Functions####################################
def GetWikipediaPage():
    global page
    if (neighborhood_level == 'place') or (neighborhood_level == 'county subdivision') or (neighborhood_level == 'county'): #Don't bother looking for wikipedia page if zip code
        try:
            wikipedia_page_search_term    = (neighborhood + ', ' + hood_state)
            page                          =  wikipedia.page(wikipedia_page_search_term)
                
        except Exception as e:
            print(e,': problem getting wikipedia page')
            
    elif (neighborhood_level == 'custom'):
        try:
            wikipedia_page_search_term    = (neighborhood + ', ' + comparison_area )
            page                          =  wikipedia.page(wikipedia_page_search_term)        
        except Exception as e:
            print(e,': problem getting wikipedia page')

def GetWalkScore(lat,lon):

    lat = str(lat)
    lon = str(lon)
    url = """https://api.walkscore.com/score?format=json&address=None&""" + """lat=""" + lat + """&lon=""" + lon + """&transit=1&bike=1&wsapikey=""" + walkscore_api_key
    print('Getting Walk Score: ', url)
   
    
    walkscore_response = requests.get(url).json()
    # print(walkscore_response)
    
    #Get Walk score from response
    try:
        walk_score           = walkscore_response['walkscore']
        walk_description     = walkscore_response['description']
        walk_table_entry     = ('Walk Score: ' + str(walk_score) + ' (' + walk_description + ')')
    except Exception as e:
        print(e,'could not get walk score')
        walk_table_entry     = 'NA'
    
    #Get Transit score from response
    try:
        transit_score        = walkscore_response['transit']['score']
        transit_description  = walkscore_response['transit']['description']
        transit_table_entry  = ('Transit Score: ' + str(transit_score) + ' (' + transit_description + ')')

    except Exception as e:
        print(e,'could not get transit score')
        transit_table_entry  = 'Transit Score: NA'


    #Get Bike score from response
    try:
         bike_score           =  walkscore_response['bike']['score']
         bike_description     =  walkscore_response['bike']['description']
         bike_table_entry     = ('Bike Score: ' + str(bike_score)  + ' (' + bike_description + ')')

    except Exception as e:
        print(e,'could not get bike score')
        bike_table_entry     = 'Bike Score: NA'

    
    #Return a list of the 3 scores
    walk_scores = [walk_table_entry, transit_table_entry, bike_table_entry]
    return(walk_scores)

def GetYelpData(lat,lon,radius):
    print('Getting Yelp Data')
    #Return a dictionary where each key is a business caategory and the values are a list of the 5 most recomended businesseses on Yelp.com
    business_categories = {'retail':[], 'banks, gyms':[], 'parks and recreation':[], 'education':[], 'transportation':[]}

    try:
        for category in business_categories.keys():
            response              = yelp_api.search_query(categories=category, longitude = lon, latitude = lat, radius = radius,sort_by = 'best_match')
            
            #Loop through the results of the yelp search and pull business names
            for i in range(5):
                business_name = response['businesses'][i]['name']
                business_categories[category].append(business_name)
        
            time.sleep(1)
    except Exception as e:
        print(e)
        
        
    return(business_categories)
    

    # bar_response              = yelp_api.search_query(categories='bars', longitude=lon, latitude=lat, radius = radius,sort_by = 'distance') # , limit=5)
    # closest_bar               = bar_response['businesses'][0]['name']
    # closest_bar_distance      = bar_response['businesses'][0]['distance']
    # pprint(bar_response)
    # print('The closest bar from the subject property on Yelp.com is ' + str(closest_bar) + ' which is ' + str(closest_bar_distance) + ' meters from the subjet property.')
    

    # number_bar_search_results = bar_response['total']
    # print('There are ' + str(number_bar_search_results) + ' bars within ' + str(radius) + ' meters of the subjet property based on a search of Yelp.com')
    

    
    # restaurants_response             = yelp_api.search_query(categories='restaurant', longitude=lon, latitude=lat, radius = radius,sort_by = 'distance') # , limit=5)
    # # pprint(restaurants_response)
    # closest_restaurant               = restaurants_response['businesses'][0]['name']
    # closest_restaurant_distance      = restaurants_response['businesses'][0]['distance']
    # print('The closest restaurant from the subject property on Yelp.com is ' + str(closest_restaurant) + ' which is ' + str(closest_restaurant_distance) + ' meters from the subjet property.')
    
    # number_restaurant_search_results = restaurants_response['total']
    # print('There are ' + str(number_restaurant_search_results) + ' restaurants within ' + str(radius) + ' meters of the subjet property based on a search of Yelp.com')
    
def FindNearestAirport(lat,lon):
    
    #Specify the file path to the airports shape file
    airport_map_location = os.path.join(data_location,'Airports','Airports.shp')
    
    
    #Open the shapefile
    airport_map = shapefile.Reader(airport_map_location)
   

    #Loop through each feature/point in the shape file
    
    for i in range(len(airport_map)):
        airport        =  airport_map.shape(i)
        airport_record = airport_map.shapeRecord(i)
        
        if airport_record.record['Fac_Type'] != 'AIRPORT':
            continue


        airport_coord = airport.points
        dist = mpu.haversine_distance( (airport_coord[0][1], airport_coord[0][0]), (lat, lon)) #measure distance between airport and subject property   
        # print(dist)

        if i == 0:
            min_dist           = dist
            cloest_airport_num = i
        elif i > 0 and dist < min_dist:
            min_dist           = dist
            cloest_airport_num = i

    closest_airport = airport_map.shapeRecord(cloest_airport_num)
    airport_lang = ('The closest airport to the geographic center of ' + neighborhood + ' is ' + closest_airport.record['Fac_Name'].title() + ' which is an ' +  closest_airport.record['Fac_Type'].lower() + ' in ' + closest_airport.record['City'].title() + ', ' + closest_airport.record['State_Name'].title() + '.' )
    airport_lang = airport_lang.replace('Intl','International')
    return(airport_lang)

def FindNearestHighways(lat,lon):
    
    #Specify the file path to the airports shape file
    road_map_location = os.path.join(data_location,'North_American_Roads','North_American_Roads.shp')
    
    #Open the shapefile
    road_map    = shapefile.Reader(road_map_location)
    # print(road_map.bbox)

    #Loop through each feature in the shape file
    for i in range(len(road_map)):
        road        =  road_map.shape(i)
        road_record = road_map.shapeRecord(i)
        

        road_coord = road.points
        # print(road_coord)
        # fish
        try:
            dist = mpu.haversine_distance( (road_coord[0][1], road_coord[0][0]), (lat, lon)) #measure distance between airport and subject property   
        except:
            dist = dist
        # print(dist)

        if i == 0:
            min_dist           = dist
            cloest_road_num = i
        elif i > 0 and dist < min_dist:
            min_dist           = dist
            cloest_road_num = i

    closest_road = road_map.shapeRecord(cloest_road_num)
    return('The closest road to the geographic center of ' + neighborhood + ' is '+  closest_road.record['ROADNAME'].title() + ' which is a ' +  str(closest_road.record['LANES']) + ' lane ' +  closest_road.record['ADMIN'].lower() + ' highway' + ' with a speed limit of ' + str(closest_road.record['SPEEDLIM']) + '.'  )

def SearchGreatSchoolDotOrg():
    print('Getting education data')
    if os.path.exists(os.path.join(hood_folder_map,'education_map.png')): #If we already have a map for this area skip it 
        return()
   
    

    try:

        if neighborhood_level == 'custom':
            search_term = (neighborhood + ', ' + comparison_area)
        else:
            search_term = (neighborhood + ', ' + hood_state)

        #Search https://www.greatschools.org/ for the area
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        browser = webdriver.Chrome(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop','chromedriver.exe')),options=options)
        browser.get('https://www.greatschools.org/')
        
        #Write hood name in box
        Place = browser.find_element_by_class_name("search_form_field")
        Place.send_keys(search_term)
        time.sleep(1.5)
        
        #Submit hood name for search
        Submit = browser.find_element_by_class_name('search_form_button')
        Submit.click()
        time.sleep(3)
        
        #Zoom out map
        pyautogui.moveTo(3261, y=1045)
        time.sleep(1)
        for i in range(1):
           pyautogui.click()
           time.sleep(1)


        time.sleep(3)


        if 'Leahy' in os.environ['USERPROFILE']: #differnet machines have different screen coordinates
            print('Using Mikes coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(1167,872, 2049, 1316) ) #left, top, width, and height
        
        elif 'Dominic' in os.environ['USERPROFILE']:
            print('Using Doms coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(3680,254,1968 ,1231) ) #left, top, width, and height
        
        else:
            im2 = pyautogui.screenshot(region=(1167,872, 2049, 1316) ) #left, top, width, and height
        
        time.sleep(.25)
        im2.save(os.path.join(hood_folder_map,'education_map.png'))
        im2.close()
        time.sleep(1)
        browser.quit()
    except Exception as e:
        print(e)
        try:
            browser.quit()
        except:
            pass

def ApartmentDotComSearchTerm():
    #Takes the name of the city or neighborhood and creates a url for apartments.com
    if neighborhood_level == 'place':
        search_term = 'https://www.apartments.com/' + '-'.join(neighborhood.lower().split(' ')) + '-' + hood_state.lower() + '/'
    elif neighborhood_level == 'custom':
        search_term = 'https://www.apartments.com/' + '-'.join(neighborhood.lower().split(' ')) + '-' + '-'.join(comparison_area.lower().split(' ')) +  '-' + hood_state.lower() + '/'


    return(search_term)

def ApartmentsDotComSearch():
    print('Seraching Apartments.com:',ApartmentDotComSearchTerm())
    try:
        search_term = ApartmentDotComSearchTerm() 

        response    = requests.get(search_term,
                                   headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.106 Safari/537.36"}
                                   )
        soup_data   = BeautifulSoup(response.text, 'html.parser')
        
        marketing_blurb_section = soup_data.find(id='marketingBlurb')
        marketing_paragraphs    = marketing_blurb_section.find_all('p')
        
        descriptive_paragraphs = []
        for count,paragraph in enumerate(marketing_paragraphs):
            if ('Learn'  in paragraph.text) and ('More' in paragraph.text):
                return(descriptive_paragraphs)
            if count > 3 :
                continue
            if 'Let Apartments.com help you find' in  paragraph.text:
                continue          
            descriptive_paragraphs.append(paragraph.text)
        
        return(descriptive_paragraphs)
    
    
    except Exception as e:
        print(e)
        return([''])

def Zoneomics(address):
    #Searches the Zoneomics API for screenshot of Local Zoning
    print('Getting Zoneomics Zoning Map')
    url = 'https://www.zoneomics.com/api/get_zone_screen_shot?address=address&api_key=api_key&map_zoom_level=17'

    #    data = {
    #    'key': zoneomics_api_key,
    #    'address': address
    #        }

    #     try:
    #         response = requests.get(url, params=data).json()
    #     except Exception as e:
    #         print(e)
    #         response = [{}]

#Main data function
def GetData():
    #List of 5 Year American Community Survey Variables here: https://api.census.gov/data/2019/acs/acs5/variables.html
    #List of 2010 Census variables here:                      https://api.census.gov/data/2010/dec/sf1/variables.html
    #List of 2020 Redistricting variables here:               https://api.census.gov/data/2020/dec/pl/variables.html
    print('Getting Data')
    # global total_number_households, average_household_size
    global overview_table_data
    global neighborhood_household_size_distribution,comparison_household_size_distribution
    global neighborhood_tenure_distribution, comparison_tenure_distribution
    global neighborhood_time_to_work_distribution, comparison_time_to_work_distribution
    global neighborhood_method_to_work_distribution
    global neighborhood_age_data,comparison_age_data
    global neighborhood_housing_value_data,comparison_housing_value_data
    global neighborhood_number_units_data,comparison_number_units_data
    global neighborhood_household_income_data, comparison_household_income_data
    global neighborhood_top_occupations_data,comparison_top_occupations_data
    global neighborhood_year_built_data, comparison_year_built_data   
    global walk_score_data
    global nyc_community_district

    print('Getting Data for ' + neighborhood)

    neighborhood_household_size_distribution          = GetHouseholdSizeData(     geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Neighborhood households by size
    neighborhood_tenure_distribution                  = GetHousingTenureData(     geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Tenure (owner occupied/renter)
    neighborhood_housing_value_data                   = GetHousingValues(         geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Owner Occupied housing units by value
    neighborhood_year_built_data                      = GetHouseYearBuiltData(    geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Units by year structure built
    neighborhood_method_to_work_distribution          = GetTravelMethodData(      geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Travel Mode to Work
    neighborhood_household_income_data                = GetHouseholdIncomeValues( geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Households by household income data
    neighborhood_time_to_work_distribution            = GetTravelTimeData(        geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Travel Time to Work
    neighborhood_number_units_data                    = GetNumberUnitsData(       geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Units by units in building
    neighborhood_age_data                             = GetAgeData(               geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Population by age data
    neighborhood_top_occupations_data                 = GetTopOccupationsData(    geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Top Employment Occupations
    
    print('Getting Data For ' + comparison_area)
    comparison_household_size_distribution            = GetHouseholdSizeData(    geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_tenure_distribution                    = GetHousingTenureData(    geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_housing_value_data                     = GetHousingValues(        geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')    
    comparison_year_built_data                        = GetHouseYearBuiltData(   geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_household_income_data                  = GetHouseholdIncomeValues(geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')   
    comparison_time_to_work_distribution              = GetTravelTimeData(       geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    
    comparison_age_data                               = GetAgeData(              geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_number_units_data                      = GetNumberUnitsData(      geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')    
    comparison_top_occupations_data                   = GetTopOccupationsData(   geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    
    #Walk score
    walk_score_data                                   = GetWalkScore(            lat = latitude, lon = longitude                                                    )

    #Overview Table Data
    overview_table_data                               = GetOverviewTable(hood_geographic_level = neighborhood_level ,comparison_geographic_level = comparison_level)
    nyc_community_district                            = DetermineNYCCommunityDistrict(lat = latitude, lon = longitude )
    
#####################################################Graph Related Functions####################################
def SetGraphFormatVariables():
    global graph_width, graph_height, scale,tickfont_size,left_margin,right_margin,top_margin,bottom_margin,legend_position,paper_backgroundcolor,title_position
    global fig_width

    #Set graph size and format variables
    marginInches = 1/18
    ppi = 96.85 
    width_inches = 6.5
    height_inches = 3.3
    fig_width                     = 4.5 #width for the pngs (graph images) we insert into report document


    graph_width  = (width_inches - marginInches)   * ppi
    graph_height = (height_inches  - marginInches) * ppi

    #Set scale for resolution 1 = no change, > 1 increases resolution. Very important for run time of main script. 
    scale = 3

    #Set tick font size (also controls legend font size)
    tickfont_size = 8 

    #Set Margin parameters/legend location
    left_margin   = 0
    right_margin  = 0
    top_margin    = 75
    bottom_margin = 10
    legend_position = 1.10

    #Paper color
    paper_backgroundcolor = 'white'

    #Title Position
    title_position = .95    

def CreateHouseholdSizeHistogram():
    print('Creating Household size graph')

    household_size_categories = ['1','2','3','4','5','6','7+']
    fig = make_subplots(specs=[[{"secondary_y": False}]])
    
    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_household_size_distribution,
           x=household_size_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )
    fig.add_trace(
    go.Bar(y=comparison_household_size_distribution,
           x=household_size_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Households by Household Size",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'household_size_graph.png'),engine='kaleido',scale=scale)

def CreateHouseholdTenureHistogram():
    print('Creating Household tenure graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    tenure_categories = ['Renter Occupied','Owner Occupied']
    
    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_tenure_distribution,
           x=tenure_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )
    fig.add_trace(
    go.Bar(y=comparison_tenure_distribution,
           x=tenure_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Occupied Housing Units by Tenure",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'household_tenure_graph.png'),engine='kaleido',scale=scale)

def CreateHouseholdNumberUnitsInBuildingHistogram():
    print('Creating Household by number of units in structure graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    number_units_categories = ['Single Family Homes','Townhomes','Duplexes','3-4 Units','5-9 Units','10-19 Units','20-49 Units','50 >= Units']
   

    #Add Bars with neighborhood distribution
    fig.add_trace(
    go.Bar(y=neighborhood_number_units_data,
           x=number_units_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    #Add Bars with comparison distribution
    fig.add_trace(
    go.Bar(y=comparison_number_units_data,
           x=number_units_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Housing Units by Units in Structure",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'household_units_in_structure_graph.png'),engine='kaleido',scale=scale)

def CreateHouseholdYearBuiltHistogram():
    print('Creating Household Year Built graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    year_built_categories = ['2014 >=','2010-2013','2000-2009','1990-1999','1980-1989','1970-1979','1960-1969','1950-1959','1940-1949','<= 1939']
    year_built_categories.reverse()

    #Add Bars with neighborhood year built data
    fig.add_trace(
    go.Bar(y=neighborhood_year_built_data,
           x=year_built_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    #Add bars for comparison area
    fig.add_trace(
    go.Bar(y=comparison_year_built_data,
           x=year_built_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Housing Units by Year Structure Built",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'household_year_built_graph.png'),engine='kaleido',scale=scale)

def CreateHouseholdValueHistogram():
    print('Creating Household value graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    housing_value_categories = ['$10,000 <','$10,000-14,999','$15,000-19,999','$20,000-24,999','$25,000-29,999','$30,000-34,000','$35,000-39,999','$40,000-49,000','$50,000-59,9999','$60,000-69,999','$70,000-79,999','$80,000-89,999','$90,000-99,999','$100,000-124,999','$125,000-149,999','$150,000-174,999','$175,000-199,999','$200,000-249,999','$250,000-299,999','$300,000-399,999','$400,000-499,999','$500,000-749,999','$750,000-999,999','$1,000,000-1,499,999','$1,500,000-1,999,999','$2,000,000 >=']
    assert len(neighborhood_housing_value_data) == len(comparison_housing_value_data)
    assert len(housing_value_categories) == len(neighborhood_housing_value_data) == len(comparison_housing_value_data)
    #Add Bars with neighborhood house value distribution
    fig.add_trace(
    go.Bar(y=neighborhood_housing_value_data,
           x=housing_value_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    fig.add_trace(
    go.Bar(y=comparison_housing_value_data,
           x=housing_value_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Owner Occupied Housing Units by Value",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_xaxes(tickangle = 45, tickfont = dict(size=tickfont_size-1))       
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'household_value_graph.png'),engine='kaleido',scale=scale)

def CreatePopulationByAgeHistogram():
    print('Creating Population by Age Graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    age_ranges = ['0-19','20-24','25-34','35-49','50-66','67+']
    
    assert len(neighborhood_age_data) == len(age_ranges) 

    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_age_data,
           x=age_ranges,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    #Add bars with comparison area age distribution
    fig.add_trace(
    go.Bar(y=comparison_age_data,
           x=age_ranges,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Population by Age",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )


    fig.update_xaxes(tickangle = 0)  
    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'population_by_age_graph.png'),engine='kaleido',scale=scale)

def CreatePopulationByIncomeHistogram():
    print('Creating Population by Income Graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    income_categories = ['< $10,000',
                         '$10,000-14,999',
                         '$15,000-19,999',
                         '$20,000-24,999',
                         '$25,000-29,999',
                         '$30,000-34,999',
                         '$35,000-39,999',
                         '$40,000-44,999',
                         '$45,000-49,999',
                         '$50,000-59,999',
                         '$60,000-74,999',
                         '$75,000-99,999',
                         '$100,000-124,999',
                         '$125,000-149,999',
                         '$150,000-199,999',
                         '> $200,000']

    assert len(income_categories) == len(neighborhood_household_income_data) == len(comparison_household_income_data)
    
    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_household_income_data,
           x=income_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    #Add bars for comparison area        
    fig.add_trace(
    go.Bar(y=comparison_household_income_data,
           x=income_categories,
           name=comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )
    
    
    #Set Title
    fig.update_layout(
    title_text="Households by Household Income",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_xaxes(tickangle = 45)       
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'population_by_income_graph.png'),engine='kaleido',scale=scale)

def CreateTopOccupationsHistogram():
    print('Creating Top Occupations Graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])
    
    occupations_categories       =  ['Management and Business','Service','Sales and Office','Natural Resources','Production']
    neighborhood_top_occupations = neighborhood_top_occupations_data
    assert len(occupations_categories) == len(neighborhood_top_occupations)
    
    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_top_occupations,
           x=occupations_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )

    
    
    #Set Title
    fig.update_layout(
    title_text="Top Employment Occupations",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'top_occupations_graph.png'),engine='kaleido',scale=scale)

def CreateTravelTimeHistogram():
    print('Creating Travel Time to work Graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    travel_time_categories = ['< 5 Minutes','5-9 Minutes','10-14 Minutes','15-19 Minutes','20-24 Minutes','25-29 Minutes','30-34 Minutes','35-39 Minutes','40-44 Minutes','45-59 Minutes','60-89 Minutes','> 90 Minutes']
    #Add Bars with neighborhood household size distribution
    fig.add_trace(
    go.Bar(y=neighborhood_time_to_work_distribution,
           x = travel_time_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )
    fig.add_trace(
    go.Bar(y=comparison_time_to_work_distribution,
           x=travel_time_categories,
           name = comparison_area,
           marker_color="#B3C3FF")
            ,secondary_y=False
            )


    
    #Set Title
    fig.update_layout(
    title_text="Travel Time to Work",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    # #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'travel_time_graph.png'),engine='kaleido',scale=scale)

def CreateTravelModeHistogram():
    print('Creating Travel Mode to work Graph')
    fig = make_subplots(specs=[[{"secondary_y": False}]])

    travel_method_categories = ['Drove Alone','Car Pooled','Public Transportation','Walked','Worked from Home','Biked','Other']
    assert len(neighborhood_method_to_work_distribution) == len(travel_method_categories)
    fig.add_trace(
    go.Bar(y=neighborhood_method_to_work_distribution,
           x=travel_method_categories,
           name=neighborhood,
           marker_color="#4160D3")
            ,secondary_y=False
            )
    
    #Set Title
    fig.update_layout(
    title_text="Travel Mode to Work",    

    title={
        'y':title_position,
        'x':0.5,
        'xanchor': 'center',
        'yanchor': 'top'},
                    
                    )
    
    #Set Legend Layout
    fig.update_layout(
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=legend_position ,
        xanchor="center",
        x=0.5,
        font_size = tickfont_size
                )

                      )
    
    fig.update_yaxes(title=None)

    #Set Font and Colors
    fig.update_layout(
    font_family="Avenir Next LT Pro",
    font_color='#262626',
    font_size = 10.5,
    paper_bgcolor=paper_backgroundcolor,
    plot_bgcolor ="White"
                     )

    #Set size and margin
    fig.update_layout(
    margin=dict(l=left_margin, r=right_margin, t=top_margin, b= bottom_margin),
    height    = graph_height,
    width     = graph_width,
        
                    )



    #Add % to  axis ticks
    fig.update_yaxes(ticksuffix = '%', tickfont = dict(size=tickfont_size),tickformat='.1f',secondary_y=False)       
    fig.write_image(os.path.join(hood_folder,'travel_mode_graph.png'),engine='kaleido',scale=scale)

def CreateGraphs():
    print('Creating Graphs')
    try:
        CreateHouseholdSizeHistogram()
    except Exception as e:
        print(e,'unable to create household size graph')
    
    try:
        CreateHouseholdTenureHistogram()
    except Exception as e:
        print(e,'unable to create household tenure graph')
    
    try:
        CreateHouseholdValueHistogram()
    
    except Exception as e:
        print(e,'unable to create housing value graph')
    
    try:
        CreateHouseholdYearBuiltHistogram()

    except Exception as e:
        print(e,'unable to create year built graph')
    
    try:
        CreateHouseholdNumberUnitsInBuildingHistogram()
    except Exception as e:
        print(e,'unable to create number units graph')
    
    try:
        CreatePopulationByAgeHistogram()
    except Exception as e:
        print(e,'unable to create population by age graph')

    try:
        CreatePopulationByIncomeHistogram()
    except Exception as e:
        print(e,'unable to create population by income graph')

    try:
        CreateTopOccupationsHistogram()
    except Exception as e:
        print(e,'unable to create top occupation graph')

    try:
        CreateTravelTimeHistogram()
    except Exception as e:
        print(e,'unable to create travel time graph')

    try:
        CreateTravelModeHistogram()
    except Exception as e:
        print(e,'unable to create travel mode graph')

#####################################################Language Related Functions####################################
def WikipediaTransitLanguage(category):
    #Searches through a wikipedia page for a number of section titles and returns the text from them (if any)
    try:
        wikipedia_search_terms_df = pd.read_csv(os.path.join(project_location,'Data','General Data','Wikipedia Transit Related Search Terms.csv'))
        wikipedia_search_terms_df = wikipedia_search_terms_df.loc[wikipedia_search_terms_df['category'] == category]
        

        language = [] 
        for search_term in wikipedia_search_terms_df['search term']:
            section = page.section(search_term)
            # print(search_term)
            # print(section)
            
            if section != None:
                language.append(section)
      
        # print(language)
        if language != []:
            return(' '.join(language))

        else:
            if category == 'car':
                return(neighborhood + ' is not connected by any major highways or roads.')

            elif category == 'bus':
                return(neighborhood + ' does not have public bus service.')

            elif category == 'air':
                #If nothing on wikipedia, use this function to look for more information
                return(FindNearestAirport(lat = latitude, lon = longitude))
              

            elif category == 'train':
                return(neighborhood + ' is not served by any commuter or light-rail lines.')
            else:
                return('')

    except Exception as e:
        print(e,'problem getting wikipedia language for ' + category)
        return('')

def SummaryLangauge():
    print('Creating Summary Langauge')
    try:
        wikipedia_summary = (wikipedia.summary((neighborhood + ',' + hood_state)))
    except Exception as e:
        print(e)
        wikipedia_summary = ('')

    try:
        apartmentsdotcomlanguage = ApartmentsDotComSearch() #neighborhood summary pulled from Apartments.com
    except Exception as e:
        print(e)
        apartmentsdotcomlanguage = ('')
        
    return[wikipedia_summary,apartmentsdotcomlanguage]

def CommunityAssetsLanguage():
    print('Creating Community Assets Langauge')
    community_assets_language = (neighborhood + ' has a number of community assets. ')
    return([community_assets_language])

def CarLanguage():
    print('Creating auto Langauge')
    wikipedia_car_language     = WikipediaTransitLanguage(category='car')
    
    if wikipedia_car_language == '':
        return(FindNearestHighways(lat = latitude, lon = longitude))
    else:
        return(wikipedia_car_language)

def PlaneLanguage():
    print('Creating plane Langauge')

    wikipedia_plane_language     = WikipediaTransitLanguage(category='air')
    nearest_airport_language     = FindNearestAirport(lat = latitude, lon = longitude)

    if wikipedia_plane_language == '':
        return(nearest_airport_language)
    else:
        return(wikipedia_plane_language)

def BusLanguage():
    print('Creating bus Langauge')

    wikipedia_bus_language = WikipediaTransitLanguage(category='bus')
    return(wikipedia_bus_language)

def TrainLanguage():
    print('Creating train Langauge')
    wikipedia_train_language = WikipediaTransitLanguage(category='train')
    return(wikipedia_train_language)

def OutlookLanguage():
    print('Creating Outlook Langauge')
    pop_growth_description = '[negative/modest/moderate/strong]'

    outlook_language = (neighborhood + 
                        ' is a '     + 
                        hood_place_type + 
                        ' in '          + 
                        comparison_area + 
                        ', '            + 
                        comparison_state_full_name + 
                        ' well-served by [interstate highways, public transportation, and recreational amenities]. ' +
                        
                        #Growth sentance
                        'It has seen ' +
                        pop_growth_description +
                        ' population growth over the past decade, a trend that is expected to continue in the near-term.'
                        
                         )
    
    return([outlook_language])

    # return('Neighborhood analysis can best be summarized by referring to neighborhood life cycles. ' +
    #       'Neighborhoods are perceived to go through four cycles, the first being growth, the second being stability, the third decline, and the fourth revitalization. ' +
    #       'It is our observation that the subject’s neighborhood is exhibiting several stages of the economic life, with an overall predominance of stability and both limited decline and limited revitalization in some sectors. ' +
    #       'The immediate area surrounding the subject, has had a historically low vacancy level and is located just to the south of the ------ submarket,' +
    #       """ which has multiple office and retail projects completed within the past two years and more development in the subject’s immediate vicinity either under construction or preparing to break ground."""+
    #       ' The proximity of the ________ and ________ will ensure the neighborhood will continue ' +
    #       'to attract growth in the long-term.')
    
def TransportationOverviewLanguage():
    print('Creating Transporation Overview Langauge')

    try:
        transportation_language         =  page.section('Transportation')
    except Exception as e:
        print(e)
        transportation_language         = ''
    
    return(transportation_language)

def HousingIntroLanguage():
	print('Creating housing intro Langauge')
	housing_intro_language = ('Housing is one of the most identifiable characteristics of an area. Different elements related to housing, such as the property type, ' +
		' renter/owner mix, housing age, and household characteristics play crucial roles in how an area is defined. ' +
		'In ' + neighborhood + ', housing is diverse, with a variety of types, tenure status, age, and price points. ')
    
	return([housing_intro_language])

def HousingTypeTenureLanguage():
    print('Creating housing type and tenure Langauge')
    number_units_categories = ['Single Family Homes','Townhomes','Duplexes','3-4 Units','5-9 Units','10-19 Units','20-49 Units','50 >= Units']
    most_common_category = number_units_categories[neighborhood_number_units_data.index(max(neighborhood_number_units_data))] #get the most common number units category
    #second_most_common_category = number_units_categories[neighborhood_number_units_data.index(max-1( #get 2nd most common house type
    hood_owner_occupied_fraction        =  neighborhood_tenure_distribution[1] 
    comparsion_owner_occupied_fraction  =  comparison_tenure_distribution[1]

    if hood_owner_occupied_fraction > comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'higher than'
        own_or_rent = 'the majority of households own instead of rent.'
    elif hood_owner_occupied_fraction < comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'lower than'
        own_or_rent = 'the majority of households rent instead of own.'
    elif hood_owner_occupied_fraction > comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'equal to'
        own_or_rent = 'an equal share of households rent or own.'        
    else:
        hood_owner_ouccupied_higher_lower   =  '[lower than/higher than/equal to]'
    
    housing_type_tenure_langugage = ('Data from the the most recent American Community Survey indicates a presence of single family homes, some smaller multifamily properties, along with larger garden style properties, and even some buildings with 50+ units. ' +
        most_common_category + ' are the most common form of housing in ' + neighborhood + ', followed by ' + '#second_most_common_category' + '. According to the most recent American Community Survey, ' +
            "{:,.1f}%".format(hood_owner_occupied_fraction)                        +   
           ' of the housing units in '                                             + 
           neighborhood                                                            + 
           ' were occupied by their owner. '                                       +
           'This percentage of owner-occupation is '                                +
           hood_owner_ouccupied_higher_lower                                        + 
           ' the '                                                                  +
           comparison_area                                                          +
           ' level of '                                                             +
           "{:,.1f}%".format(comparsion_owner_occupied_fraction)                    +
           '.'
           )

    return([housing_type_tenure_langugage])

def HousingValueLanguage():
    print('Creating Household by value Langauge')

    housing_value_categories = ['$10,000 <','$10,000-14,999','$15,000-19,999','$20,000-24,999','$25,000-29,999','$30,000-34,000','$35,000-39,999','$40,000-49,000','$50,000-59,9999','$60,000-69,999','$70,000-79,999','$80,000-89,999','$90,000-99,999','$100,000-124,999','$125,000-149,999','$150,000-174,999','$175,000-199,999','$200,000-249,999','$250,000-299,999','$300,000-399,999','$400,000-499,999','$500,000-749,999','$750,000-999,999','$1,000,000-1,499,999','$1,500,000-1,999,999','$2,000,000 >=']
    assert len(neighborhood_housing_value_data) == len(housing_value_categories) == len(comparison_housing_value_data)
    #Estimate a median household income from a category freqeuncy distribution
    hood_median_value_range     = FindMedianCategory(frequency_list = neighborhood_housing_value_data, category_list = housing_value_categories)
    hood_median_value_range     = hood_median_value_range.replace('$','')
    hood_median_value_range     = hood_median_value_range.replace(',','').split('-')
    hood_median_value           = round((int(hood_median_value_range[0]) + int(hood_median_value_range[1]))/2,1)

    #Estimate a median household income from a category freqeuncy distribution
    comp_median_value_range     = FindMedianCategory(frequency_list = comparison_housing_value_data, category_list = housing_value_categories)
    comp_median_value_range     = comp_median_value_range.replace('$','')
    comp_median_value_range     = comp_median_value_range.replace(',','').split('-')
    comp_median_value           = round((int(comp_median_value_range[0]) + int(comp_median_value_range[1]))/2,1)
    
    hood_largest_value_category = housing_value_categories[neighborhood_housing_value_data.index(max(neighborhood_housing_value_data))] #get the most common income category
    comp_largest_value_category = housing_value_categories[comparison_housing_value_data.index(max(comparison_housing_value_data))]

    value_language = (  'Homes in '                                           +
                       neighborhood                                           + 
                       ' have a median value of about '                       + 
                        "${:,.0f}".format(hood_median_value)                  +
                       ', compared to '                                       +
                       "${:,.0f}".format(comp_median_value)                   +
                       ' for '                                            +  
                       comparison_area                                        +
                       '. In '                                                + 
                       neighborhood                                           + 
                       ', the largest share of homes have a value between ' +
                       hood_largest_value_category +
                       ', compared to ' +
                       comp_largest_value_category        +
                       ' for '           +
                        comparison_area +
                        '.'
                        )
    
    return([value_language])

def HousingYearBuiltLanguage():
    print('Creating House by Year Built Langauge')
    
    year_built_categories       = ['2014','2010-2013','2000-2009','1990-1999','1980-1989','1970-1979','1960-1969','1950-1959','1940-1949','1939']
    year_built_categories.reverse()

    #Median Year Built for hoodS
    hood_median_yrblt_range     =  FindMedianCategory(frequency_list = neighborhood_year_built_data, category_list = year_built_categories)
    
    if len(hood_median_yrblt_range) == 4:
        hood_median_yrblt = int(hood_median_yrblt_range)
    else:
        hood_median_yrblt_range     = hood_median_yrblt_range.split('-')
        hood_median_yrblt           = round((int(hood_median_yrblt_range[0]) + int(hood_median_yrblt_range[1]))/2,1)
    

    #Median Year Built for comparison area
    comp_median_yrblt_range     =  FindMedianCategory(frequency_list = comparison_year_built_data, category_list = year_built_categories)
    
    if len(comp_median_yrblt_range) == 4:
        comp_median_yrblt = int(comp_median_yrblt_range)
    else:
        comp_median_yrblt_range     = comp_median_yrblt_range.split('-')
        comp_median_yrblt           = round((int(comp_median_yrblt_range[0]) + int(comp_median_yrblt_range[1]))/2,1)
    

    #Largest cateorgies for hood and comparison area
    hood_largest_yrblt_category = year_built_categories[neighborhood_year_built_data.index(max(neighborhood_year_built_data))] #get the most common income category
    comp_largest_yrblt_category = year_built_categories[comparison_year_built_data.index(max(comparison_year_built_data))]


    yrblt_language = (  'Homes in '                                         +
                       neighborhood                                         + 
                       ' have a median year built of about '                + 
                        "{:.0f}".format(hood_median_yrblt)                  +
                       ', compared to '                                     +
                        "{:.0f}".format(comp_median_yrblt)                  +
                        ' for '                                             +
                        comparison_area                                     +
                       '. '                                                 +
                       
                       'In '                                                + 
                       neighborhood                                         + 
                       ', the largest share of homes were built between '   +
                       hood_largest_yrblt_category                          +
                       ', compared to '                                     +
                       comp_largest_yrblt_category                          +
                       ' for '                                              +
                        comparison_area                                     +
                        '.'
                        )
    
    return([yrblt_language])

def EmploymentLanguage():
    print('Creating Employment by Industry angauge')
    return(['The majority of working age residents are employed in the ______, ______, and _______ industries. ' + 'Given its large share of employees in the production industry, '])

def HouseholdSizeLanguage():
    print('Creating Household by Size Langauge')

    household_size_categories = ['1','2','3','4','5','6','7+']


    #Median Household size for hood
    hood_median_size   = int(FindMedianCategory(frequency_list = neighborhood_household_size_distribution, category_list = household_size_categories).replace('+',''))
    comp_median_size   = int(FindMedianCategory(frequency_list = comparison_household_size_distribution,   category_list = household_size_categories).replace('+',''))
    
    #Largest cateogy for hood and comparsion area
    hood_largest_time_category = household_size_categories[neighborhood_household_size_distribution.index(max(neighborhood_household_size_distribution))] #get the most common household size category
    comp_largest_time_category = household_size_categories[comparison_household_size_distribution.index(max(comparison_household_size_distribution))]

    household_size_language = ('Households in '                                        +
                               neighborhood                                            + 
                              ' have a median size of '                                + 
                              "{:,.0f} people".format(hood_median_size)                +
                              '. '                                                     +

                              'In '                                                    + 
                              neighborhood                                             + 
                              ', the largest share of households have '                +
                              hood_largest_time_category                               +
                              ' people, compared to '                                  +
                              comp_largest_time_category                               +
                              ' for '                                                  +
                              comparison_area                                          +
                              '.'
                            )
    
    return([household_size_language])

def PopulationAgeLanguage():
    print('Creating Population by Age Langauge')
    age_ranges = ['0-19','20-24','25-34','35-49','50-66','67+']

    hood_median_age_range      = FindMedianCategory(frequency_list = neighborhood_age_data, category_list = age_ranges)
    hood_median_age_range      = hood_median_age_range.replace(',','').split('-')
    hood_median_age            = round((int(hood_median_age_range[0]) + int(hood_median_age_range[1]))/2,1)
    
    hood_largest_age_category  = age_ranges[neighborhood_age_data.index(max(neighborhood_age_data))] #get the most common income category
    comp_largest_age_category  = age_ranges[comparison_age_data.index(max(comparison_age_data))]

    age_language = ('The median age in '                                                        +
                       neighborhood                                                         + 
                       ' is around '                         + 
                        "{:,.0f}".format(hood_median_age)                                   +
                       '. '                                   +
                       
                       'In '                                                                + 
                       neighborhood                                                         + 
                       ', the largest age range is between ' +
                       hood_largest_age_category                                            +
                       ', compared to '                                                     +
                       comp_largest_age_category                                            +
                       ' for '                                                              +
                        comparison_area                                                     +
                        '.'
                    )

    
    return([age_language])

def IncomeLanguage():
    print('Creating HH Income Langauge')
    income_categories = ['< $10,000',
                         '$10,000-14,999',
                         '$15,000-19,999',
                         '$20,000-24,999',
                         '$25,000-29,999',
                         '$30,000-34,999',
                         '$35,000-39,999',
                         '$40,000-44,999',
                         '$45,000-49,999',
                         '$50,000-59,999',
                         '$60,000-74,999',
                         '$75,000-99,999',
                         '$100,000-124,999',
                         '$125,000-149,999',
                         '$150,000-199,999',
                         '> $200,000']

    #Estimate a median household income from a category freqeuncy distribution
    hood_median_income_range   = FindMedianCategory(frequency_list=neighborhood_household_income_data, category_list=income_categories)
    hood_median_income_range   = hood_median_income_range.replace('$','')
    hood_median_income_range   = hood_median_income_range.replace(',','').split('-')
    hood_median_income         = round((int(hood_median_income_range[0]) + int(hood_median_income_range[1]))/2,1)
    

    hood_largest_income_category = income_categories[neighborhood_household_income_data.index(max(neighborhood_household_income_data))] #get the most common income category
    

    #If not the last or first category
    if (hood_largest_income_category != income_categories[len(income_categories)-1]) and (hood_largest_income_category != income_categories[0]) :
        hood_largest_income_category = 'between ' + hood_largest_income_category

    elif (hood_largest_income_category == income_categories[0]):
        hood_largest_income_category = 'under ' + hood_largest_income_category.replace('<','')
    
    elif (hood_largest_income_category == income_categories[len(income_categories)-1]):
        hood_largest_income_category = 'over ' + hood_largest_income_category.replace('>','')



    comp_largest_income_category = income_categories[comparison_household_income_data.index(max(comparison_household_income_data))]

    #If not the last or first category
    if (comp_largest_income_category != income_categories[len(income_categories)-1]) and (comp_largest_income_category != income_categories[0]) :
        comp_largest_income_category = 'between ' + comp_largest_income_category

    elif (comp_largest_income_category == income_categories[0]):
        comp_largest_income_category = 'under ' + comp_largest_income_category.replace('>','')
    
    elif (comp_largest_income_category == income_categories[len(income_categories)-1]):
        comp_largest_income_category = 'over ' + comp_largest_income_category.replace('>','')


    income_language = ('Households in '                                      +
                       neighborhood                                          + 
                       ' have a median household income of around '          + 
                        "${:,.0f}".format(hood_median_income)                 +
                       '. '                    +
                       
                       'In '                                                 + 
                       neighborhood                                          + 
                       ', the largest share of households have a household income ' +
                       hood_largest_income_category +
                       ', compared to ' +
                       comp_largest_income_category        +
                       ' for '           +
                        comparison_area +
                        '.'
                        )
    
    return([income_language])

def TravelMethodLanguage():
    print('Creating Travel Method Langauge')
    travel_method_categories = ['driving alone','car pooling','public transportation','walking','working from home','biking','other']
    hood_largest_travel_category      = travel_method_categories[neighborhood_method_to_work_distribution.index(max(neighborhood_method_to_work_distribution))] #get the most common income category
    hood_largest_travel_category_frac = neighborhood_method_to_work_distribution[neighborhood_method_to_work_distribution.index(max(neighborhood_method_to_work_distribution))]

    travel_method_language = ('In ' + neighborhood + ', the most common method for traveling to work is ' + hood_largest_travel_category.lower()  + ' with ' +  "{:,.1f}%".format(hood_largest_travel_category_frac) + ' of commuters using it.')
    return([travel_method_language])
    
def TravelTimeLanguage():
    print('Creating Travel Time Langauge')
    travel_time_categories = ['< 5 Minutes','5-9 Minutes','10-14 Minutes','15-19 Minutes','20-24 Minutes','25-29 Minutes','30-34 Minutes','35-39 Minutes','40-44 Minutes','45-59 Minutes','60-89 Minutes','> 90 Minutes']


    #Estimate a median household income from a category freqeuncy distribution
    hood_median_time_range   = FindMedianCategory(frequency_list=neighborhood_time_to_work_distribution, category_list = travel_time_categories) 
    hood_median_time_range   = hood_median_time_range.replace(' Minutes','')
    hood_median_time_range   = hood_median_time_range.replace(',','').split('-')
    hood_median_time         = (int(hood_median_time_range[0]) + int(hood_median_time_range[1]))/2

    #Estimate a median household income from a category freqeuncy distribution
    comp_median_time_range   = FindMedianCategory(frequency_list=comparison_time_to_work_distribution, category_list = travel_time_categories) 
    comp_median_time_range   = comp_median_time_range.replace(' Minutes','')
    comp_median_time_range   = comp_median_time_range.replace(',','').split('-')
    comp_median_time         = (int(comp_median_time_range[0]) + int(comp_median_time_range[1]))/2
    
    hood_largest_time_category = travel_time_categories[neighborhood_time_to_work_distribution.index(max(neighborhood_time_to_work_distribution))] #get the most common income category
    comp_largest_time_category = travel_time_categories[comparison_time_to_work_distribution.index(max(comparison_time_to_work_distribution))]

    time_language = ('Commuters in '                                        +
                       neighborhood                                           + 
                       ' have a median commute time of about '                      + 
                        "{:,.0f} minutes".format(hood_median_time)                   +
                       '. '                     +
                       
                       'In '                                                  + 
                       neighborhood                                           + 
                       ', the largest share of commuters have a commute between ' +
                       hood_largest_time_category +
                       ', compared to ' +
                       comp_largest_time_category        +
                       ' for '           +
                        comparison_area +
                        '.'
                        )
    
    return([time_language])

def LocationIQPOIListLanguage(lat,lon,category):
    #Searches the Locate IQ API for points of interest
    print('Searching Location IQ API for: ',category)

    url = "https://us1.locationiq.com/v1/nearby.php"

    data = {
    'key': location_iq_api_key,
    'lat': lat,
    'lon': lon,
    'tag': category,
    'radius': 5000,
    'format': 'json'
        }

    try:
        response = requests.get(url, params=data).json()
    except Exception as e:
        return('')
    
    try:
        poi_list = (neighborhood + ' has several ' + category + ' assets including: ')
        for poi in response:
            try:
                poi_name      = poi['name']
                # print(poi_name)
                poi_type      = poi['type']
                poi_city      = poi['address']['city']
                poi_sentence  = (' ' + poi_name + ', ' )
                poi_list = poi_list + poi_sentence
            except:
                continue
            #For cities/towns, restrict points of interest to those inside the city limits

            # if neighborhood_level == 'place':
            #     if neighborhood == poi_city:
            #         poi_sentence = poi_list + poi_sentence
            # else:
        time.sleep(.25)
        return(poi_list)
    except Exception as e:
            print(e)

def CreateLanguage():
    
    print('Creating Langauge')

    global bus_language,car_language,plane_language,train_language,transportation_language,summary_langauge,conclusion_langauge,community_assets_language
    global yelp_language
    global airport_language
    global apartmentsdotcomlanguage
    global housing_type_tenure_language
    global employment_language
    global population_age_language
    global income_language
    global travel_method_language, travel_time_language
    global housing_value_language,year_built_language
    global household_size_language
    global bank_language, food_language, hospital_language, park_language, retail_language, edu_language
    
    summary_langauge                   =  SummaryLangauge()
    transportation_language            =  TransportationOverviewLanguage()
    
    try:
        community_assets_language          =  CommunityAssetsLanguage()
    except Exception as e:
        print(e,'unable to get community assets langauge')
        community_assets_language = []
    
    try:
        housing_type_tenure_language       = HousingTypeTenureLanguage()
    except Exception as e:
        housing_type_tenure_language       = []
    
    try:
        housing_value_language             = HousingValueLanguage()
    except Exception as e:
        housing_value_language             = []
    
    try:
        year_built_language                = HousingYearBuiltLanguage()
    except Exception as e:
        year_built_language                = []


    #Communtiy assets langauge variables
    bank_language                      = LocationIQPOIListLanguage(lat = latitude, lon = longitude , category = 'bank' ) 
    food_language                      = LocationIQPOIListLanguage(lat = latitude, lon = longitude,  category = 'food' ) 
    hospital_language                  = LocationIQPOIListLanguage(lat = latitude, lon = longitude,  category = 'hospital' ) 
    park_language                      = LocationIQPOIListLanguage(lat = latitude, lon = longitude,  category = 'park' ) 
    retail_language                    = LocationIQPOIListLanguage(lat = latitude, lon = longitude,  category = 'retail' ) 
    edu_language                       = LocationIQPOIListLanguage(lat = latitude, lon = longitude,  category = 'school' ) 


    #Paragraph Language
    employment_language                = EmploymentLanguage()
    population_age_language            = PopulationAgeLanguage()
    income_language                    = IncomeLanguage()
    household_size_language            = HouseholdSizeLanguage()
    travel_method_language             = TravelMethodLanguage() 
    travel_time_language               = TravelTimeLanguage()

    #Transit Table Language
    bus_language                       = BusLanguage() 
    train_language                     = TrainLanguage()
    car_language                       = CarLanguage()
    plane_language                     = PlaneLanguage()

    conclusion_langauge                = OutlookLanguage()

#####################################################Report document related functions####################################
def SetPageMargins(document,margin_size):
    sections = document.sections
    for section in sections:
        section.top_margin    = Inches(margin_size)
        section.bottom_margin = Inches(margin_size)
        section.left_margin   = Inches(margin_size)
        section.right_margin  = Inches(margin_size)

def SetDocumentStyle(document):
    style = document.styles['Normal']
    font = style.font
    font.name = 'Avenir Next LT Pro (Body)'
    font.size = Pt(9)

def AddTitle(document):
    main_title = document.add_heading('Neighborhood & Demographic Overview',level=0) 
    main_title.style = document.styles['Heading 1']
    main_title.paragraph_format.space_after  = Pt(6)
    main_title.paragraph_format.space_before = Pt(12)
    main_title_style = main_title.style
    main_title_style.font.name = "Avenir Next LT Pro Light"
    main_title_style.font.size = Pt(18)
    main_title_style.font.bold = False
    main_title_style.font.color.rgb = RGBColor.from_string('3F65AB')
    main_title_style.element.xml
    rFonts = main_title_style.element.rPr.rFonts
    rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro Light")

    # glance_paragraph                               = document.add_paragraph(neighborhood + ' at a Glance')
    # glance_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
    # glance_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

def AddHeading(document,title,heading_level,heading_number,font_size): #Function we use to insert the headers other than the title header
            heading = document.add_heading(title,level=heading_level)
            heading.style = document.styles[heading_number]
            heading_style =  heading.style
            heading_style.font.name = "Avenir Next LT Pro Light"
            heading_style.font.size = Pt(font_size)
            heading_style.font.bold = False
            heading.paragraph_format.space_after  = Pt(6)
            heading.paragraph_format.space_before = Pt(12)

            #Color
            heading_style.font.color.rgb = RGBColor.from_string('3F65AB')            
            heading_style.element.xml
            rFonts = heading_style.element.rPr.rFonts
            rFonts.set(qn("w:asciiTheme"), "Avenir Next LT Pro")

def AddTableTitle(document,title):
    table_title_paragraph = document.add_paragraph(title)
    table_title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_title_paragraph.paragraph_format.space_after  = Pt(6)
    table_title_paragraph.paragraph_format.space_before = Pt(12)
    for run in table_title_paragraph.runs:
                    font = run.font
                    font.name = 'Avenir Next LT Pro Medium'

def Citation(document,text):
    citation_paragraph = document.add_paragraph()
    citation_paragraph.paragraph_format.space_after  = Pt(0)
    citation_paragraph.paragraph_format.space_before = Pt(0)
    run = citation_paragraph.add_run('Source: ' + text)
    font = run.font
    font.name = primary_font
    font.size = Pt(8)
    font.italic = True
    font.color.rgb  = RGBColor.from_string('929292')
    citation_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if text != 'Google Maps':
        blank_paragraph = document.add_paragraph('')
        blank_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def GetMap():
    print('Getting Map')
    try:
        #Search Google Maps for hood
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        browser = webdriver.Chrome(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop','chromedriver.exe')),options=options)
        browser.get('https:google.com/maps')
            
        #Write hood name in box
        Place = browser.find_element_by_class_name("tactile-searchbox-input")

        if neighborhood_level != 'custom':
            Place.send_keys((neighborhood + ', ' + hood_state))
        
        elif neighborhood_level == 'custom':
            Place.send_keys((neighborhood + ', ' + comparison_area))

        #Submit hood name for search
        Submit = browser.find_element_by_class_name('nhb85d-BIqFsb')
        Submit.click()
        time.sleep(12)

        # first photo, up close and personal. no zoom needed
        if 'Leahy' in os.environ['USERPROFILE']: #differnet machines have different screen coordinates
            print('Using Mikes coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(1358,465, 2142, 1404) ) #left, top, width, and height
        
        elif 'Dominic' in os.environ['USERPROFILE']:
            print('Using Doms coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(3680,254,1968 ,1231) ) #left, top, width, and height
        
        else:
            im2 = pyautogui.screenshot(region=(1089,276, 2405, 1754) ) #left, top, width, and height
        time.sleep(1)
        im2.save(os.path.join(hood_folder_map,'map.png'))
        time.sleep(3)

        # second photo, zoomed out
        zoomout = browser.find_element_by_xpath("""//*[@id="widget-zoom-out"]/div""")
        for i in range(3):
            zoomout.click() 
        time.sleep(3)


        if 'Leahy' in os.environ['USERPROFILE']: #differnet machines have different screen coordinates
            print('Using Mikes coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(1358,465, 2142, 1404) ) #left, top, width, and height
        
        elif 'Dominic' in os.environ['USERPROFILE']:
            print('Using Doms coordinates for screenshot')
            im2 = pyautogui.screenshot(region=(3680,254,1968 ,1231) ) #left, top, width, and height
        
        else:
            im2 = pyautogui.screenshot(region=(1089,276, 2405, 1754) ) #left, top, width, and height
        time.sleep(5)

        im2.save(os.path.join(hood_folder_map,'map2.png'))
        im2.close()

        #Wait till we have saved both png images before proceeding
        while (os.path.exists(os.path.join(hood_folder_map,'map.png')) == False) or (os.path.exists(os.path.join(hood_folder_map,'map2.png')) == False):
            pass

        browser.quit()
    except Exception as e:
         print(e)
         try:
            browser.quit()
         except:
            pass

def add_border(input_image, output_image, border):
    #adds border around png images
    img = Image.open(input_image)
    if isinstance(border, int) or isinstance(border, tuple):
        bimg = ImageOps.expand(img, border=border)
    else:
        raise RuntimeError('Border is not an image or tuple')
    bimg.save(output_image)

def OverlayMapImages():
    print("Creating overlayed map image")
    map_path  =  os.path.join(hood_folder_map,'map.png')
    map2_path = os.path.join(hood_folder_map,'map2.png')
    map3_path = os.path.join(hood_folder_map,'map3.png')
    
    #Make sure we have map 1 and map 2 in order to create map 3 (the overlayed map image)
    try:
        assert (os.path.exists(map_path)) and (os.path.exists(map2_path))
    except:
        print('Unable to make overlayed map')
        return()

    #Open zommed out map
    img1 = Image.open(map2_path)
    
    add_border(map_path, output_image = map_path, border=5)
    time.sleep(.2)
    #Open zommed in map
    img2 = Image.open(map_path)

    #Reduce size of zommed in image by a constant factor
    image_reduction_scale = 3
    img2 = img2.resize((int(img2.size[0]/image_reduction_scale),int(img2.size[1]/image_reduction_scale)))
    
    #Add the zoomed in map on top of the zoomed out map and save as new png image
    # No transparency mask specified,                                      
    # simulating an raster overlay
    img1.paste(img2, (img1.size[1] - 25,900))
    
    img1.save(map3_path)
    
def AddMap(document):
    map_path        = os.path.join(hood_folder_map,'map.png')
    map2_path       = os.path.join(hood_folder_map,'map2.png')
    map3_path       = os.path.join(hood_folder_map,'map3.png')
    nyc_cd_map_path = os.path.join(nyc_cd_map_location,nyc_community_district,'map.png')
    
    if (os.path.exists(map_path) == False) or (os.path.exists(map2_path) == False): #If we don't have a zommed in map image or a zoomed out map, create one
        GetMap()    
    if os.path.exists(map3_path) == False: #If we don't have an image with a zommed in map overlayed on zoomed out map, create one
        OverlayMapImages()
   
    print('Adding Map') 
    if os.path.exists(map3_path):
        paragraph = document.add_paragraph('')
        paragraph.add_run().add_picture(map3_path,width=Inches(6.5))
        paragraph.paragraph_format.space_after         = Pt(0)
        
    
    if os.path.exists(nyc_cd_map_path):
        print('Adding NYC Community District Map') 
        nyc_map = document.add_picture(nyc_cd_map_path,width=Inches(6.5))

def PageBreak(document):
    #Add page break
    page_break_paragraph = document.add_paragraph('')
    run = page_break_paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)

def AddDocumentParagraph(document,language_variable):
    assert type(language_variable) == list
    for paragraph in language_variable:
        if paragraph == '':
            continue
        par                                               = document.add_paragraph(paragraph)
        par.alignment                                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        par.paragraph_format.space_after                  = Pt(primary_space_after_paragraph)
        summary_format                                    = document.styles['Normal'].paragraph_format
        summary_format.line_spacing_rule                  = WD_LINE_SPACING.SINGLE
        style = document.styles['Normal']
        font = style.font
        font.name = 'Avenir Next LT Pro Light'
        par.style = document.styles['Normal']

def AddTable(document,data_for_table): #Function we use to insert our overview table into the report document
    #list of list where each list is a row for our table
     
    #make sure each list inside the list of lists has the same number of elements
    for row in data_for_table:
        for row2 in data_for_table:
            assert len(row) == len(row2)


    #create table object
    tab = document.add_table(rows=len(data_for_table), cols=len(data_for_table[0]))
    tab.alignment     = WD_TABLE_ALIGNMENT.CENTER
    tab.allow_autofit = True
    #loop through the rows in the table
    for current_row ,(row,row_data_list) in enumerate(zip(tab.rows,data_for_table)): 

    
        row.height = Inches(0)

        #loop through all cells in the current row
        for current_column,(cell,cell_data) in enumerate(zip(row.cells,row_data_list)):
            cell.text = str(cell_data)

            if current_row == 0:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM


            #set column widths
            if current_column == 0:
                cell.width = Inches(1.25)

            elif current_column == 1:
                cell.width = Inches(1.25)




            #add border to top row
            if current_row == 1:
                    tcPr = cell._element.tcPr
                    tcBorders = OxmlElement("w:tcBorders")
                    top = OxmlElement('w:top')
                    top.set(qn('w:val'), 'single')
                    tcBorders.append(top)
                    tcPr.append(tcBorders)
            
            #loop through the paragraphs in the cell and set font and style
            for paragraph in cell.paragraphs:
                if current_column > 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                     paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                #Make first row flush agaisnt bottom
                if current_row == 0:
                    paragraph.paragraph_format.space_after   = Pt(0)
                    paragraph.paragraph_format.space_before  = Pt(0)


                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(8)
                    run.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    #make first row bold
                    if current_row == 0: 
                        font.bold = True
                        font.name = 'Avenir Next LT Pro Demi'

def AddTwoColumnTable(document,pic_list,lang_list):
    #Insert the transit graphics(car, bus,plane, train)
    tab = document.add_table(rows=0, cols=2)
    for pic,lang in zip(pic_list,lang_list):
        pic_path = os.path.join(graphics_location,pic)
        if os.path.exists(pic_path) == False:
            continue
        row_cells = tab.add_row().cells
        
        left_paragraph = row_cells[0].paragraphs[0]
        run            = left_paragraph.add_run()
        if pic == 'car.png':
            run.add_text(' ')
        run.add_picture(pic_path)

        right_paragraph = row_cells[1].paragraphs[0]
        run             = right_paragraph.add_run()
        run.add_text(str(lang))

    #We have now defined our table object,loop through all rows then all cells in each current row
    for row in tab.rows:
        for current_column,cell in enumerate(row.cells):
            #Set Width for cell
            if current_column == 0:
                cell.width = Inches(.2)
            elif current_column == 1:
                cell.width = Inches(6)

def AddPointOfInterestsTable(document,data_for_table): #Function we use to insert our table with Location IQ points of interest into the report document
    print(data_for_table)
    print(type(data_for_table))
    print(type(data_for_table[0]))

    #Convert the data from location IQ from json to list of list where each list is a row for the table
    converted_data_for_table = [ list(data_for_table[0].keys())  ]
    
    
    for i in data_for_table:
        new_list = list(i.values())
        converted_data_for_table.append(new_list)

    print(converted_data_for_table)
    assert type(converted_data_for_table) == list

    #make sure each list inside the list of lists has the same number of elements
    for row in converted_data_for_table:
        for row2 in converted_data_for_table:
            assert len(row) == len(row2)


    #create table object
    tab = document.add_table(rows=len(converted_data_for_table), cols=len(converted_data_for_table[0]))
    tab.alignment     = WD_TABLE_ALIGNMENT.CENTER
    tab.allow_autofit = True
    #loop through the rows in the table
    for current_row ,(row,row_data_list) in enumerate(zip(tab.rows,converted_data_for_table)): 

    
        row.height = Inches(0)

        #loop through all cells in the current row
        for current_column,(cell,cell_data) in enumerate(zip(row.cells,row_data_list)):
            cell.text = str(cell_data)

            if current_row == 0:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM


            #set column widths
            if current_column == 0:
                cell.width = Inches(1.25)

            elif current_column == 1:
                cell.width = Inches(1.25)




            #add border to top row
            if current_row == 1:
                    tcPr = cell._element.tcPr
                    tcBorders = OxmlElement("w:tcBorders")
                    top = OxmlElement('w:top')
                    top.set(qn('w:val'), 'single')
                    tcBorders.append(top)
                    tcPr.append(tcBorders)
            
            #loop through the paragraphs in the cell and set font and style
            for paragraph in cell.paragraphs:
                if current_column > 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                     paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(8)
                    run.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    #make first row bold
                    if current_row == 0: 
                        font.bold = True
                        font.name = 'Avenir Next LT Pro Demi'

def IntroSection(document):
    print('Writing Intro Section')
    AddTitle(document = document)
    AddMap(document = document)
    Citation(document,'Google Maps')
    AddHeading(document = document, title =  (neighborhood + ' at a Glance'),            heading_level = 1,heading_number='Heading 3',font_size=11)
   
    #Add neighborhood overview language
    AddDocumentParagraph(document = document,language_variable =  summary_langauge)
    AddTableTitle(document = document, title = 'Population and Household Growth')
    
    try:
        #Add Overview Table
        AddTable(document = document,data_for_table = overview_table_data )
    except Exception as e:
        print(e,'Unable to add overview table')

def CommunityAssetsSection(document):
    print('Writing Community Assets Section')
    #Community Assets Section
    AddHeading(document = document, title = 'Community Assets',            heading_level = 1,heading_number='Heading 3',font_size=11)

    AddDocumentParagraph(document = document,language_variable =  community_assets_language)

    #Table Title
    AddTableTitle(document = document, title = 'Community Assets')

    #Add Community Assets Table                 
    AddTwoColumnTable(document,pic_list      = ['bank.png','food.png','medical.png','park.png','retail.png','school.png'],lang_list =[bank_language, food_language, hospital_language, park_language, retail_language,edu_language] )

def HousingSection(document):
    print('Writing Neighborhood Section')
    AddHeading(document = document, title = 'Housing',                  heading_level = 1,heading_number='Heading 3',font_size=11)
    
    #Insert household units by units in_structure graph
    if os.path.exists(os.path.join(hood_folder,'household_units_in_structure_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_units_in_structure_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Add tenure language
    AddDocumentParagraph(document = document,language_variable =  housing_type_tenure_language)

    #Insert Household Tenure graph
    if os.path.exists(os.path.join(hood_folder,'household_tenure_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_tenure_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Add housing value language
    AddDocumentParagraph(document = document,language_variable =  housing_value_language)

    #Insert Household value graph
    if os.path.exists(os.path.join(hood_folder,'household_value_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_value_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Add language
    AddDocumentParagraph(document = document,language_variable =  year_built_language)

    #Insert household units by year built graph
    if os.path.exists(os.path.join(hood_folder,'household_year_built_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_year_built_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

def DevelopmentSection(document):
    print('Writing Development Section')
    
    #Development subsection
    AddHeading(document = document, title = 'Development',                  heading_level = 1,heading_number='Heading 3',font_size=11)

def EducationSection(document):
    print('Writing Education Section')

    AddHeading(document = document, title = 'Education',                  heading_level = 1,heading_number='Heading 3',font_size=11)

    if os.path.exists(os.path.join(hood_folder_map,'education_map.png')):
        fig = document.add_picture(os.path.join(hood_folder_map,'education_map.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'greatschools.org')

def PopulationSection(document):
    print('Writing Population Section')
    
    AddHeading(document = document, title = 'Population',                                     heading_level = 1,heading_number='Heading 3',font_size=11)
    
    #household size langauge
    AddDocumentParagraph(document = document,language_variable =  household_size_language)

    #Insert Household size graph
    if os.path.exists(os.path.join(hood_folder,'household_size_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_size_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Age langauge
    AddDocumentParagraph(document = document,language_variable =  population_age_language)
    
    #Insert population by age graph
    if os.path.exists(os.path.join(hood_folder,'population_by_age_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'population_by_age_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Income langauge
    AddDocumentParagraph(document = document,language_variable =  income_language)

    #Insert populatin by income graph
    if os.path.exists(os.path.join(hood_folder,'population_by_income_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'population_by_income_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

def EmploymentSection(document):
    print('Writing Employment Section')

    #Employment and Transportation Section
    AddHeading(document = document, title = 'Employment',                  heading_level = 1,heading_number='Heading 3',font_size=11)

    AddDocumentParagraph(document = document,language_variable =  employment_language)
    
    #Insert top occupations graph
    if os.path.exists(os.path.join(hood_folder,'top_occupations_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'top_occupations_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
        
def TransportationSection(document):
    print('Writing Transportation Section')
    #Employment and Transportation Section
    AddHeading(document = document, title = 'Transportation',                  heading_level = 1,heading_number='Heading 3',font_size=11)

    #Travel time Lanaguage
    AddDocumentParagraph(document = document,language_variable =  travel_time_language)

    #Insert Travel Time to Work graph
    if os.path.exists(os.path.join(hood_folder,'travel_time_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'travel_time_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Travel method Lanaguage
    AddDocumentParagraph(document = document,language_variable =  travel_method_language)

    #Insert Transport Method to Work graph
    if os.path.exists(os.path.join(hood_folder,'travel_mode_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'travel_mode_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Transportation Methods table
    AddTableTitle(document = document, title = 'Transportation Methods')
    AddTwoColumnTable(document,pic_list      = ['car.png','train.png','bus.png','plane.png','walk.png'],lang_list =[car_language, train_language, bus_language, plane_language,walk_score_data[0]] )

def OutlookSection(document):
    print('Writing Outlook Section')
    AddHeading(document = document, title = 'Conclusion',            heading_level = 1,heading_number='Heading 3',font_size=11)

    AddDocumentParagraph(document = document,language_variable =  conclusion_langauge)
    
def WriteReport():
    print('Writing Report')
    #Create Document
    document = Document()
    SetPageMargins(           document  = document, margin_size=1)
    SetDocumentStyle(         document = document)
    IntroSection(             document = document)
    CommunityAssetsSection(   document = document)
    HousingSection(           document = document)
    # DevelopmentSection(       document = document)
    # EducationSection(         document = document)
    PopulationSection(        document = document)
    EmploymentSection(        document = document)
    TransportationSection(    document = document)
    OutlookSection(           document = document)


    #Save report
    document.save(report_path)  

def CleanUpPNGs():
    print('Deleting PNG files')
    #Report writing done, delete figures
    files = os.listdir(hood_folder)
    for image in files:
        if image.endswith(".png"):
            os.remove(os.path.join(hood_folder, image))

def CreateDirectoryCSV():
    global service_api_csv_name
    print('Creating CSV with file path information on all existing hood reports')
    dropbox_links                  = []
    dropbox_research_names         = []
    dropbox_neighborhoods          = []
    dropbox_analysis_types         = []
    dropbox_states                 = []
    dropbox_versions               = []
    dropbox_statuses               = []
    dropbox_document_names         = []


    for (dirpath, dirnames, filenames) in os.walk(main_output_location):
        if filenames == []:
            continue
        else:
            for file in filenames:
                
                full_path = dirpath + '/' + file

                if file == 'Dropbox Neighborhoods.csv' or ('Archive' in dirpath) or ('~' in file):
                    continue
                
                if (os.path.exists(full_path.replace('_draft','_FINAL'))) and ('_draft' in full_path) or ('docx' not in full_path):
                    continue

                dropbox_document_names.append(file)
                dropbox_analysis_types.append('Neighborhood')
                dropbox_link = dirpath.replace(dropbox_root,r'https://www.dropbox.com/home')
                dropbox_link = dropbox_link.replace("\\",r'/')    
                dropbox_links.append(dropbox_link)
                dropbox_versions.append(file[0:4])
                if '_draft' in file:
                    file_status = 'Draft'
                else:
                    file_status = 'Final'

                dropbox_statuses.append(file_status)

                
                state_name    = file[5:7]
                
                try:
                    hood_name     = file.split(' - ')[1].strip()
                    research_name = state_name + ' - ' + file.split(' - ')[1].strip()
                
                except:
                    hood_name     = 'FIX FILE FORMAT'
                    research_name = 'FIX FILE FORMAT'
                
                dropbox_neighborhoods.append(hood_name)
                dropbox_research_names.append(research_name)
                dropbox_states.append(state_name)
            
        

    dropbox_df = pd.DataFrame({'Market Research Name':dropbox_research_names,
                            'Neighborhood':dropbox_neighborhoods,
                           'Analysis Type': dropbox_analysis_types,
                           'State':         dropbox_states,
                           "Dropbox Links":dropbox_links,
                           'Version':dropbox_versions,
                           'Status':dropbox_statuses,
                           'Document Name': dropbox_document_names})
    dropbox_df = dropbox_df.sort_values(by=['State','Market Research Name'])

    csv_name = 'Dropbox Neighborhoods.csv'
    service_api_csv_name = f'Dropbox Neighborhoods-{datetime.now().timestamp()}.csv'
    dropbox_df.to_csv(os.path.join(main_output_location, csv_name),index=False)

    if main_output_location == os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood'):
        dropbox_df.to_csv(os.path.join(main_output_location, service_api_csv_name),index=False)

def DecideIfWritingReport():
    global report_creation
    #Give the user 10 seconds to decide if writing reports for metro areas or individual county entries
    try:
        # report_creation = input_with_timeout('Create new report? y/n', 10).strip()
        report_creation = 'y'

    except TimeoutExpired:
        report_creation = ''

def UserSelectsNeighborhoodLevel():
    global analysis_type_number
    analysis_type_number = input('What is the geographic level of the neighborhood and comparison area?' + '\n'
  
                                '1.) = Place  vs. County'+ '\n' #+
                                '2.) = County Subdivison vs. County' + '\n' +
                                '3.) = Custom vs. Place'  + '\n' +
                                '4.) = Place  vs. County Subdivison'+ '\n' +
                                # '5.) = Zip    vs. Place'+ '\n' #+

                                # '6.) = Tract vs. Place'   + '\n' +
                                # '7.) = Tract vs. County ' + '\n' +
                                # '8.) = Tract vs. Zip'     + '\n' +
                                # '9.) = Tract vs. County Subdivison'+ '\n' +
                                # '10.) = Tract vs. Custom'+ '\n' +
                                # '11.) = Tract vs. None'+ '\n' +

                                # '12.) = Place  vs. Zip'+ '\n' +
                                # '13.) = Place  vs. Custom'+ '\n' +
                                # '14.) = Place  vs. Tract'+ '\n' +
                                # '15.) = Place  vs. None'+ '\n' +


                                # '16.) = County  vs. Place'+ '\n' +
                                # '17.) = County  vs. Tract' + '\n' +
                                # '18.) = County vs. Zip' + '\n' +
                                # '19.) = County vs. Custom'+ '\n' +
                                # '20.) = County vs. County Subdivison'+ '\n' +
                                # '21.) = County  vs. None'+ '\n' +

                                # '22.) = Zip vs. Tract '+ '\n' +
                                # '23.) = Zip vs. Custom'+ '\n' +
                                # '24.) = Zip vs. County Subdivison'+ '\n' +
                                # '25.) = Zip vs. County'+ '\n' +
                                # '26.) = Zip vs. None'+ '\n' +

                                # '27.) = County Subdivison vs. Place'  + '\n' +
                                # '28.) = County Subdivison vs. Custom' + '\n' +
                                # '29.) = County Subdivison vs. Zip'+ '\n' +
                                # '30.) = County Subdivison vs. Tract'+ '\n' +
                                # '31.) = County Subdivison vs. None'  + '\n' +
                            
                                # '32.) = Custom vs. Tract'  + '\n' +
                                # '33.) = Custom vs. County Subdivison' + '\n' +
                                # '34.) = Custom vs. County' + '\n' +
                                # '35.) = Custom vs. Zip'  + '\n' +
                                # '36.) = Custom  vs. None'  
                                ''
                                )

    
    return(int(analysis_type_number))

def GetUserInputs(analysis_type_number):
    global neighborhood_level,comparison_level

    
    #Each number corresponds to a different analysis level pair eg: place vs county, zip vs. place, etc
    if analysis_type_number == 1: #Place  vs. County
        neighborhood_level = 'place'
        comparison_level   = 'county'
    elif analysis_type_number == 2: #County Subdivison vs. County
        neighborhood_level = 'county subdivision'
        comparison_level   = 'county'
    elif analysis_type_number == 3: #Custom vs. Place
        neighborhood_level = 'custom'
        comparison_level   = 'place'
    elif analysis_type_number == 4: #Place vs. County Subdivison
        neighborhood_level = 'place'
        comparison_level   = 'county subdivision'
    # elif analysis_type_number == 5: #Zip vs. Place
    #     neighborhood_level = 'zip'
    #     comparison_level   = 'place'
    # elif analysis_type_number == 6: #Tract vs. Place
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 7: #Tract vs. County
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 8: #Tract vs. Zip
      #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 9: #Tract vs. County Subdivison
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 10: #Tract vs. Custom
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 11: #Tract vs. None
     #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 12: #Place  vs. Zip
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 13: #Place  vs. Custom
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 14: #Place  vs. Tract
     #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 15: #Place  vs. None
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 16: #County  vs. Place
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 17: #County  vs. Tract
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 18: #County vs. Zip
     #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 19: #County vs. Custom
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 20: #County vs. County Subdivison
    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 21: #County  vs. None
     #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 22: #Zip vs. Tract
     #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 23: #Zip vs. Custom
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 24: #Zip vs. County Subdivison
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 25: #Zip vs. County
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 26: #Zip vs. None
    #    #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 27: #County Subdivison vs. Place
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 28: #County Subdivison vs. Custom
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 29: #County Subdivison vs. Zip
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 30: #County Subdivison vs. Tract
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 31: #County Subdivison vs. None
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 32: #Custom vs. Tract
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 33 : #Custom vs. County Subdivison
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 34 : #Custom vs. County
        #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 35: #Custom vs. Zip
       #   neighborhood_level = ''
    #   comparison_level   = ''
    # elif analysis_type_number == 36: #Custom  vs. None
        #   neighborhood_level = ''
    #   comparison_level   = ''
    else:
            print('Not a supported level currently')
    


    global neighborhood, hood_tract, hood_zip, hood_place_fips, hood_place_type, hood_suvdiv_fips, hood_county_fips
    global hood_state, hood_state_fips, hood_state_full_name

    #Get User input on neighborhood/subject area
    if neighborhood_level == 'place':        #when our neighborhood is a town or city eg: East Rockaway Village, New York
        place_fips_info                 = ProcessPlaceFIPS(input('Enter the 7 digit Census Place FIPS Code'))
        hood_place_fips                 = place_fips_info[0]
        hood_state_fips                 = place_fips_info[1]
        neighborhood                    = place_fips_info[2]
        hood_state_full_name            = place_fips_info[3]
        hood_state                      = place_fips_info[4]
        hood_place_type                 = place_fips_info[5]

    elif neighborhood_level == 'county subdivision':     #when our neighborhood is county subdivison eg: Town of Hempstead, New York (A large town in Nassau County with several villages within it)
        subdivision_fips_info           = ProcessCountySubdivisionFIPS(county_subdivision_fips=input('Enter the 10 digit county subdivision FIPS Code'))
        hood_suvdiv_fips                = subdivision_fips_info[0]
        hood_county_fips                = subdivision_fips_info[1]
        neighborhood                    = subdivision_fips_info[2]
        hood_state_fips                 = subdivision_fips_info[3]
        hood_state_full_name            = subdivision_fips_info[4]
        hood_state                      = subdivision_fips_info[5]
        hood_place_type                 = subdivision_fips_info[6]

    elif neighborhood_level == 'tract':      #when our neighborhood is a census tract eg: Tract 106.01 in Manhattan
        tract_info                      = ProcessCountyTract(tract = input('Enter the 6 digit tract code for hood'), county_fips =  input('Enter the 5 digit County FIPS Code for the county the hood tract is in'))
        hood_county_fips                = tract_info[0]
        hood_tract                      = tract_info[1]
        neighborhood                    = tract_info[2]
        hood_state_full_name            = tract_info[3]
        hood_state                      = tract_info[4]
        hood_state_fips                 = tract_info[5]
        hood_place_type                 = 'census tract'
                    
    elif neighborhood_level == 'zip':      #When our neighborhood is a zip code eg: 11563
        zip_info                         = ProcessZipCode(zip_code=input('Enter the 5 digit zip code for hood'))
        hood_county_fips                 = zip_info[0]
        hood_zip                         = zip_info[1]
        neighborhood                     = zip_info[2]
        hood_state_full_name             = zip_info[3]
        hood_state                       = zip_info[4]
        hood_state_fips                  = zip_info[5]
        hood_place_type                  = 'zip code'

    elif neighborhood_level == 'county':      #When our neighborhood is a county eg Nassau County, New York
        county_fips_info                = ProcessCountyFIPS(county_fips =   input('Enter the 5 digit county FIPS Code for the hood'))
        hood_county_fips                = county_fips_info[0]
        hood_state_fips                 = county_fips_info[1]
        neighborhood                    = county_fips_info[2]
        hood_state_full_name            = county_fips_info[3]
        hood_state                      = county_fips_info[4]
        hood_place_type                 = 'county'

    elif neighborhood_level == 'custom': #When our neighborhood is a neighboorhood within a city (eg: Financial District, New York City)
        #Get name of hood
        neighborhood      = input('Enter the name of the custom neighborhood').strip()
        # hood_place_type   = 'neighborhood' #use this for batches of  all the hoods in a city
        hood_place_type                 = 'custom'


    global comparison_area, comparison_tract ,comparison_zip, comparison_place_fips, comparison_suvdiv_fips, comparison_county_fips
    global comparison_state, comparison_state_fips, comparison_state_full_name
    global comparison_place_type

    #Get user input on comparison area
    if comparison_level == 'county':          #When our comparison area is a county eg Nassau County, New York
        if neighborhood_level == 'place':
            county_fips_info                      = ProcessCountyFIPS(PlaceFIPSToCountyFIPS(hood_place_fips))
        else:
            county_fips_info                      = ProcessCountyFIPS(county_fips =   input('Enter the 5 digit county FIPS Code for the hood'))

        comparison_county_fips                = county_fips_info[0]
        comparison_state_fips                 = county_fips_info[1]
        comparison_area                       = county_fips_info[2]
        comparison_state_full_name            = county_fips_info[3]
        comparison_state                      = county_fips_info[4]
        comparison_place_type                 = 'county'

    elif comparison_level == 'place':        #when our comparison area is a town or city eg: East Rockaway Village, New York
        # place_fips_info                      = ProcessPlaceFIPS(place_fips) #use this for batches of  all the hoods in a city
        place_fips_info                      = ProcessPlaceFIPS(place_fips = input('Enter the 7 digit Census Place FIPS Code') )
        comparison_place_fips                = place_fips_info[0]
        comparison_state_fips                = place_fips_info[1]
        comparison_area                      = place_fips_info[2]
        comparison_state_full_name           = place_fips_info[3]
        comparison_state                     = place_fips_info[4]
        comparison_place_type                = place_fips_info[5]

    elif comparison_level == 'county subdivision':       #when our comparison area is county subdivison eg: Town of Hempstead, New York (A large town in Nassau County with several villages within it)
        subdivision_fips_info                 = ProcessCountySubdivisionFIPS(county_subdivision_fips=input('Enter the 10 digit county subdivision FIPS Code'))
        comparison_suvdiv_fips                = subdivision_fips_info[0]
        comparison_county_fips                = subdivision_fips_info[1]
        comparison_area                       = subdivision_fips_info[2]
        comparison_state_fips                 = subdivision_fips_info[3]
        comparison_state_full_name            = subdivision_fips_info[4]
        comparison_state                      = subdivision_fips_info[5]
        comparison_place_type                 = subdivision_fips_info[6]

    elif comparison_level == 'zip':        #When our comparison area is a zip code eg: 11563
        zip_info                               = ProcessZipCode(zip_code=input('Enter the 5 digit zip code for comparison area'))
        comparison_county_fips                 = zip_info[0]
        comparison_zip                         = zip_info[1]
        comparison_area                        = zip_info[2]
        comparison_state_full_name             = zip_info[3]
        comparison_state                       = zip_info[4]
        comparison_state_fips                  = zip_info[5]
        comparison_place_type                  = 'zip code'
      
    elif comparison_level == 'tract':        #when our comparison area is a census tract eg: Tract 106.01 in Manhattan
        tract_info                            = ProcessCountyTract(tract = input('Enter the 6 digit tract code'), county_fips =  input('Enter the 5 digit County FIPS Code for the county the hood tract is in'))
        comparison_county_fips                = tract_info[0]
        comparison_tract                      = tract_info[1]
        comparison_area                       = tract_info[2]
        comparison_state_full_name            = tract_info[3]
        comparison_state                      = tract_info[4]
        comparison_state_fips                 = tract_info[5]
        comparison_place_type                 = 'census tract'

    elif comparison_level == 'custom':   #When our comparison area is a neighboorhood within a city (eg: Financial District, New York City)
        comparison_area                       = input('Enter the name of the custom comparison area').strip()
        comparison_place_type                 = 'neighborhood'

    #Use comparison area state when doing a custom report
    if neighborhood_level == 'custom':
        hood_state                      = comparison_state

def UpdateServiceDb(report_type, csv_name, csv_path, dropbox_dir):
    if type == None:
        return
    print(f'Updating service database: {report_type}')

    try:
        url = f'http://market-research-service.bowery.link/api/v1/update/{report_type}'
        dropbox_path = f'{dropbox_dir}{csv_name}'
        payload = { 'location': dropbox_path }

        retry_strategy = Retry(
            total=3,
            status_forcelist=[400, 404, 409, 500, 503, 504],
            allowed_methods=["POST"],
            backoff_factor=5,
            raise_on_status=False
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        http = requests.Session()
        http.mount("https://", adapter)
        http.mount("http://", adapter)

        response = http.post(url, json=payload)
        response.raise_for_status()
        print('Service successfully updated')
    except HTTPError as e:
        print(f'Request failed with status code {response.status_code}')
        json_error = json.loads(response.content)
        print(json.dumps(json_error, indent=2))
        print('Service DB did not successfully update. Please run the script again after fixing the error.')
    finally:
        print(f'Deleting temporary CSV: ', csv_path)
        os.remove(csv_path)           

def Main(analysis_type_number):
    DeclareAPIKeys()
    DeclareFormattingParameters()
    DecideIfWritingReport()
   
    if report_creation == 'y':
        GetUserInputs(analysis_type_number) #user selects if they want to run report and gives input for report subject
        print('Preparing report for: ' + neighborhood + ' compared to ' + comparison_area)
        global latitude
        global longitude
        global current_year
        global neighborhood_shape
        coordinates = GetLatandLon()
        latitude    = coordinates[0] 
        longitude   = coordinates[1]
        neighborhood_shape = GetNeighborhoodShape()
        current_year = str(date.today().year)
        SetGraphFormatVariables()
        CreateDirectory()
        GetWikipediaPage()
        GetData()
        CreateGraphs()
        CreateLanguage()
        WriteReport()
        CleanUpPNGs()
        print('Report for: ---------' + neighborhood + ' compared to ' + comparison_area + ' Complete ----------------')

    
    #Crawl through directory and create CSV with all current neighborhood report documents
    CreateDirectoryCSV()

    #Post an update request to the Market Research Docs Service to update the database
    if main_output_location == os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood'): 
        UpdateServiceDb(report_type='neighborhoods', 
                    csv_name=service_api_csv_name, 
                    csv_path=os.path.join(main_output_location, service_api_csv_name),
                    dropbox_dir='https://www.dropbox.com/home/Research/Market Analysis/Neighborhood/')



#This is our main function that calls all other functions we will use
batch_mode = False
#EXPERIMENT IN PROGRESS BATCH HOOD RERPORTS
if batch_mode == True:
    for city,place_fips in zip(['New York'],['36-51000']):
        
        for  neighborhood in GetListOfNeighborhoods(city):
            try:
                Main(analysis_type_number =     3)
            except Exception as e:
                print(e,'REORT CREATION FAILED')
else:
    Main(analysis_type_number = UserSelectsNeighborhoodLevel())


print('Finished!')