#By Mike Leahy
#Started 06/30/2021
#Summary: This script creates reports on neighborhoods/cities for Bowery

import json
import math
import msvcrt
import os
import re
import shutil
import sys
import time
from ctypes import addressof
from datetime import date, datetime
from itertools import count
from pprint import pprint
from random import randrange
from tkinter.constants import E

import census_area
import docx
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
from bls_datasets import oes, qcew
from blsconnect import RequestBLS, bls_search
from bs4 import BeautifulSoup
from census import Census
from census_area import Census as CensusArea
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.table import CT_Row, CT_Tc
from docx.shared import Inches, Pt, RGBColor
from fredapi import Fred
from genericpath import exists
from numpy import true_divide
from osgeo import gdal, ogr
from PIL import Image, ImageOps
from plotly.subplots import make_subplots
from requests.adapters import HTTPAdapter
from requests.exceptions import HTTPError
from requests.packages.urllib3.util.retry import Retry
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from shapely.geometry import LineString, MultiPoint, Point, shape
from shapely.ops import nearest_points
from us import states
from walkscore import WalkScoreAPI
from wikipedia.wikipedia import random
from yelpapi import YelpAPI

#Define file paths
dropbox_root                   =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)') 
project_location               =  os.path.join(os.environ['USERPROFILE'], 'Dropbox (Bowery)','Research','Projects', 'Research Report Automation Project') 
main_output_location           =  os.path.join(project_location,'Output','Neighborhood')                   #testing
# main_output_location           =  os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood') #production
data_location                  =  os.path.join(project_location,'Data','Neighborhood Reports Data')
graphics_location              =  os.path.join(project_location,'Data','Graphics')
map_location                   =  os.path.join(project_location,'Data','Maps','Neighborhood Maps')

#Set formatting paramaters for reports
primary_font                  = 'Avenir Next LT Pro Light' 
primary_space_after_paragraph = 8

#Decide if you want to export data in excel files in the county folder
data_export                   = False
testing_mode                  = True
testing_mode                  = False


#Directory Realted Functions
def CreateDirectory():

    global report_path,hood_folder_map,hood_folder
    
    state_folder_map         = os.path.join(map_location,state)

    state_folder             = os.path.join(main_output_location,state)

    if neighborhood_level == 'custom':

        if os.path.exists(state_folder) == False:
            os.mkdir(state_folder)  
        
        if os.path.exists(state_folder_map) == False:
            os.mkdir(state_folder_map)  

        city_folder =  os.path.join(main_output_location,state,comparison_area)
        city_folder_map =  os.path.join(map_location,state,comparison_area)

        if os.path.exists(city_folder) == False:
            os.mkdir(city_folder) 
        
        if os.path.exists(city_folder_map) == False:
            os.mkdir(city_folder_map) 


        hood_folder              = os.path.join(main_output_location,state,comparison_area,neighborhood)
        hood_folder_map          = os.path.join(map_location,state,city_folder_map,neighborhood)

    else:
        hood_folder              = os.path.join(main_output_location,state,neighborhood)
        hood_folder_map          = os.path.join(map_location,state,neighborhood)


    for folder in [state_folder,hood_folder,state_folder_map,hood_folder_map]:
         if os.path.exists(folder):
            pass 
         else:
            os.mkdir(folder) 





    report_path = os.path.join(hood_folder,current_year + ' ' + state + ' - ' + neighborhood  + ' - hood' + '_draft.docx')

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

#Data Manipulation functions
def ConvertListElementsToFractionOfTotal(raw_list):
    #Convert list with raw totals into a list where each element is a fraction of the total
    total = sum(raw_list)

    converted_list = []
    for i in raw_list:
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

#Data Gathering Related Functions
def DeclareAPIKeys():
    global census_api_key,walkscore_api_key,google_maps_api_key,yelp_api_key,yelp_api,yelp_client_id,location_iq_api_key
    global c,c_area,walkscore_api
    
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

#Lat and Lon
def GetLatandLon():
    if testing_mode == False:
        # latitude  = float(input('enter the latitude for the subject property')) 
        # longitude = float(input('enter the longitude for the subject property'))

        # Look up lat and lon of area with geocoding using google maps api
        gmaps          = googlemaps.Client(key=google_maps_api_key) 
        
        if neighborhood_level == 'custom:':
            geocode_result = gmaps.geocode(address=(neighborhood + ', ' + comparison_area + ',' + state),)
        else:
            geocode_result = gmaps.geocode(address=(neighborhood + ',' + state),)
        
        latitude       = geocode_result[0]['geometry']['location']['lat']
        longitude      = geocode_result[0]['geometry']['location']['lng']
    
    elif testing_mode == True:
        latitude    = 40.652490
        longitude   = -73.658980

    return([latitude,longitude]) 

def GetNeighborhoodShape():
    try:
        #Method 1: Pull geojson from file with city name
        with open(os.path.join(data_location,'Neighborhood Shapes',comparison_area + '.geojson')) as infile: #Open a geojson file with the city as the name the name of the file with the neighborhood boundries for that city
            my_shape_geojson = json.load(infile)
            
        #Iterate through the features in the file (each feature is a negihborhood) and find the boundry of interest
        for i in range(len(my_shape_geojson['features'])):
            feature_hood_name = my_shape_geojson['features'][i]['properties']['name']
            if feature_hood_name == neighborhood:
                neighborhood_shape = my_shape_geojson['features'][i]['geometry']
    
        return(neighborhood_shape) 
             
    
    except Exception as e:
        print(e,'problem getting shape from city geojson file')
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
        
        neighborhood_shape = my_shape_geojson['features'][0]['geometry']
        return(neighborhood_shape) 

def FindZipCodeDictionary(zip_code_data_dictionary_list,zcta,state_fips):
    #This function takes a list of dictionaries, where each zip code gets its own dictionary. Takes a zip code and state fips code and finds and returns just that dictionary.
    #We need to use this, because the census api is causing an error that requires us to retrive data for all zip codes in the country
    for zcta_dictionary in  zip_code_data_dictionary_list:
    
        if zcta_dictionary['zip code tabulation area'] == zcta and zcta_dictionary['state'] == state_fips:
            return(zcta_dictionary)
        

    print('Could not find dictionary for given zip code: ', zcta )
            
#Household Size
def GetHouseholdSizeData(geographic_level,hood_or_comparison_area):
    print('Getting household size data for: ',hood_or_comparison_area)

    #Define variables we request from census api
    fields_list = ['H013002','H013003','H013004','H013005','H013006','H013007','H013008']
    
    #Speicify geographic level specific varaibles
    if geographic_level == 'place':

        try:

            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips
            
            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips
            
            neighborhood_household_size_distribution_raw = c.sf1.state_place(fields=fields_list,state_fips=state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting household size data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county':
        
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            
            neighborhood_household_size_distribution_raw = c.sf1.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting household size data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
        
            neighborhood_household_size_distribution_raw = c.sf1.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        except Exception as e:
            print(e, 'Problem getting household size data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip


            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip

        
            neighborhood_household_size_distribution_raw = c.sf1.state_zipcode(fields=fields_list,state_fips=state_fips,zcta=zcta)[0]
        except Exception as e:
            print(e, 'Problem getting household size data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract = hood_tract 
                county_fips = hood_county_fips


            elif hood_or_comparison_area == 'comparison area':
                tract = comparison_tract
                county_fips = comparison_county_fips
            
            neighborhood_household_size_distribution_raw = c.sf1.state_county_tract(fields=fields_list, state_fips = state_fips,county_fips=county_fips,tract=tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting household size data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':

        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.sf1.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        neighborhood_household_size_distribution_raw = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )
    





    #General data manipulation (same for all geographic levels)
    neighborhood_household_size_distribution = []
    for field in fields_list:
            neighborhood_household_size_distribution.append(neighborhood_household_size_distribution_raw[field])
        
    neighborhood_household_size_distribution = ConvertListElementsToFractionOfTotal(neighborhood_household_size_distribution)
    return(neighborhood_household_size_distribution)

#Household Tenure
def GetHousingTenureData(geographic_level,hood_or_comparison_area):
    #Occupied Housing Units by Tenure
    print('Getting tenure data for: ',hood_or_comparison_area)

    fields_list = ['H004004','H004003','H004002']
    if geographic_level == 'place':
        
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            neighborhood_tenure_distribution_raw    = c.sf1.state_place(fields=fields_list,state_fips=state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting household tenure data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips

            neighborhood_tenure_distribution_raw    = c.sf1.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        except Exception as e:
            print(e, 'Problem getting household tenure data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        
        if hood_or_comparison_area == 'hood':
            county_fips = hood_county_fips
            subdiv_fips = hood_suvdiv_fips

        elif hood_or_comparison_area == 'comparison area':
            county_fips = comparison_county_fips
            subdiv_fips = comparison_suvdiv_fips
        
        neighborhood_tenure_distribution_raw    = c.sf1.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
                        
    elif geographic_level == 'zip':
        
        if hood_or_comparison_area == 'hood':
            zcta = hood_zip


        elif hood_or_comparison_area == 'comparison area':
            zcta = comparison_zip
        
        neighborhood_tenure_distribution_raw    = c.sf1.state_zipcode(fields=fields_list,state_fips=state_fips,zcta=zcta)[0]

    elif geographic_level == 'tract':
        
        if hood_or_comparison_area == 'hood':
            tract = hood_tract 
            county_fips = hood_county_fips


        elif hood_or_comparison_area == 'comparison area':
            tract = comparison_tract
            county_fips = comparison_county_fips
        
        neighborhood_tenure_distribution_raw    = c.sf1.state_county_tract(fields=fields_list,state_fips=state_fips, county_fips=county_fips, tract=tract)[0]

    elif geographic_level == 'custom':
        
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.sf1.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        neighborhood_tenure_distribution_raw = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )


    neighborhood_tenure_distribution = []
    for field in fields_list:
        neighborhood_tenure_distribution.append(neighborhood_tenure_distribution_raw[field])

    neighborhood_tenure_distribution = ConvertListElementsToFractionOfTotal(neighborhood_tenure_distribution)

    #add together the owned free and clear percentage with the owned with a mortgage percentage to simply an owner-occupied fraction
    neighborhood_tenure_distribution[1] = neighborhood_tenure_distribution[1] +  neighborhood_tenure_distribution[2]
    del neighborhood_tenure_distribution[2]

    return(neighborhood_tenure_distribution)
    
#Age Related Data Functions
def GetAgeData(geographic_level,hood_or_comparison_area):
    print('Getting age breakdown for: ',hood_or_comparison_area)
    #Return a list with the fraction of the population in different age groups 

    #Define 2 lists of variables, 1 for male age groups and another for female
    male_fields_list   = ["B01001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(3,26)]  #5 Year ACS age variables for men range:  B01001_003E - B01001_025E
    female_fields_list =  ["B01001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(27,50)] #5 Year ACS age variables for women range:  B01001_027E - B01001_049E

   
    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            male_age_data = c.acs5.state_place(fields=male_fields_list, state_fips=state_fips,place=place_fips)[0]
            female_age_data = c.acs5.state_place(fields=female_fields_list,state_fips=state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            
            male_age_data = c.acs5.state_county(fields=male_fields_list,state_fips=state_fips,county_fips=county_fips)[0]
            female_age_data = c.acs5.state_county(fields=female_fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        except Exception as e:
            print(e, 'Problem getting age data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips

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

#Housing related data functions
def GetHousingValues(geographic_level,hood_or_comparison_area):
    print('Getting housing value data for: ',hood_or_comparison_area)

    #5 Year ACS household  value range:  B25075_002E -B25075_027E
    fields_list = ["B25075_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,28)]

    
    if geographic_level == 'place':

        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips
            
            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            household_value_raw_data = c.acs5.state_place(fields = fields_list, state_fips = state_fips, place = place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips

            household_value_raw_data = c.acs5.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            household_value_raw_data = c.acs5.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta =  hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta =  comparison_zip
            
            household_value_raw_data       = c.acs5.zipcode(fields = fields_list, zcta = '*')
            household_value_raw_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   household_value_raw_data  , zcta = zcta, state_fips = state_fips )
            
        except Exception as e:
            print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
        
            household_value_raw_data = c.acs5.state_county_tract(fields = fields_list, state_fips = state_fips, county_fips = county_fips, tract = tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        household_value_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )
    
    try:
        #Create an empty list and place the values from the dictionary inside of it
        household_value_data = []
        for field in fields_list:
            household_value_data.append(household_value_raw_data[field])

        household_value_data = ConvertListElementsToFractionOfTotal(household_value_data)
        
        return(household_value_data)
    except Exception as e:
        print(e, 'Problem getting housing value data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )

#Number of Housing Units based on number of units in building
def GetNumberUnitsData(geographic_level,hood_or_comparison_area):
    print('Getting housing units by number of units data for: ', hood_or_comparison_area)
    
    
    owner_occupied_fields_list  = ["B25032_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(3,11)]   #5 Year ACS owner occupied number of units variables range:  B25032_003E - B25032_010E
    renter_occupied_fields_list = ["B25032_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(14,22)]  #5 Year ACS renter occupied number of units variables range: B25032_014E - B25032_021E 

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips
        
            owner_occupied_units_raw_data = c.acs5.state_place(fields = owner_occupied_fields_list,state_fips=state_fips,place=place_fips)[0]
            renter_occupied_units_raw_data = c.acs5.state_place(fields = renter_occupied_fields_list,state_fips=state_fips,place=place_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting number units data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county':
        try:
            
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            
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

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
        
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

#Household Income data functions
def GetHouseholdIncomeValues(geographic_level,hood_or_comparison_area):
    print('Getting household income data for: ',hood_or_comparison_area)

    #5 Year ACS household income range:  B19001_002E - B19001_017E
    fields_list = ["B19001_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,18)]

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips
        

            household_income_data = c.acs5.state_place(fields=fields_list, state_fips=state_fips, place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting household income data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips

            household_income_data = c.acs5.state_county(fields=fields_list, state_fips=state_fips, county_fips=county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting household income data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            household_income_data = c.acs5.state_county_subdivision(fields = fields_list, state_fips = state_fips, county_fips = county_fips, subdiv_fips = subdiv_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting household income data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
                

            household_income_data       = c.acs5.zipcode(fields = fields_list, zcta = '*')
            household_income_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   household_income_data  , zcta = zcta, state_fips = state_fips )

        except Exception as e:
            print(e, 'Problem getting household income data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
            
            household_income_data = c.acs5.state_county_tract(fields=fields_list, state_fips=state_fips, county_fips=county_fips, tract=tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting household income data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        household_income_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )




    #Create an empty list and place the values from the dictionary inside of it
    household_income_breakdown = []
    for field in fields_list:
        household_income_breakdown.append(household_income_data[field])


    total_pop = sum(household_income_breakdown) 

    total_income_breakdown = []
    for i in household_income_breakdown:
        total_income_breakdown.append((i/total_pop) * 100)

    assert len(total_income_breakdown) == 16
    return(total_income_breakdown)

#Occupations Data
def GetTopOccupationsData(geographic_level,hood_or_comparison_area):
    print('Getting occupation data for: ',hood_or_comparison_area)

    cateogries_dict = {'B24011_002E':'Management and Business','B24011_018E':'Service','B24011_026E':'Sales and Office','B24011_029E':'Natural Resources','B24011_036E':'Production'}

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            data = c.acs5.state_place(fields=list(cateogries_dict.keys()),state_fips=state_fips,place=place_fips)[0]
            del data['state']
            del data['place']
        except Exception as e:
            print(e, 'Problem getting top occupations data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif  geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            data = c.acs5.state_county(fields=list(cateogries_dict.keys()),state_fips=state_fips,county_fips=county_fips)[0]
            del data['state']
            del data['county']
        
        except Exception as e:
            print(e, 'Problem getting top occupations data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            data = c.acs5.state_county_subdivision(fields=list(cateogries_dict.keys()),state_fips=state_fips,county_fips=county_fips, subdiv_fips = subdiv_fips)[0]
            del data['state']
            del data['county']
            del data['county subdivision']
        
        except Exception as e:
            print(e, 'Problem getting top occupations data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
        

            data       = c.acs5.zipcode(fields = list(cateogries_dict.keys()), zcta = '*')
            data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   data  , zcta = zcta, state_fips = state_fips )
            
            del data['state']
            del data['zip code tabulation area']
        
        except Exception as e:
            print(e, 'Problem getting top occupations data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
        
    elif geographic_level == 'tract':
        try:        
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips

            data = c.acs5.state_county_tract(fields=list(cateogries_dict.keys()),state_fips=state_fips,county_fips=county_fips, tract = tract)[0]
        
            del data['state']
            del data['county']
            del data['tract']
        
        except Exception as e:
            print(e, 'Problem getting top occupations data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(list(cateogries_dict.keys()), neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = cateogries_dict.keys() )
        print(data)



    
    data = dict((cateogries_dict[key], value) for (key, value) in data.items())
    data = {k: v for k, v in sorted(data.items(), key=lambda item: item[1])}

    total_workers = sum(list(data.values()))
   

    #Convert from raw ammount to percent of total
    for key in data:
        data[key] = (data.get(key)/total_workers) * 100
        
    return(data)

#Year Housing Built Data
def GetHouseYearBuiltData(geographic_level,hood_or_comparison_area):
    print('Getting year built data for: ',hood_or_comparison_area)

    #5 Year ACS household year house built range:  B25034_002E -B25034_011E
    fields_list = ["B25034_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,12)]

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips
            year_built_raw_data = c.acs5.state_place(fields=fields_list,state_fips=state_fips,place=place_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            year_built_raw_data = c.acs5.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            year_built_raw_data = c.acs5.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips = subdiv_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
             
            year_built_raw_data       = c.acs5.zipcode(fields = fields_list, zcta = '*')
            year_built_raw_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   year_built_raw_data  , zcta = zcta, state_fips = state_fips )

        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
            
            year_built_raw_data = c.acs5.state_county_tract(fields=fields_list,state_fips=state_fips,county_fips=county_fips,tract = tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        year_built_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )


    #Create an empty list and place the values from the dictionary inside of it
    year_built_data = []
    for field in fields_list:
        year_built_data.append(year_built_raw_data[field])

    #Convert list with raw totals into a list where each element is a fraction of the total
    year_built_data = ConvertListElementsToFractionOfTotal(year_built_data) 
    year_built_data.reverse()

    return(year_built_data)

#Travel Time to Work
def GetTravelTimeData(geographic_level,hood_or_comparison_area):
    print('Getting travel time data for: ', hood_or_comparison_area)
    #5 Year ACS travel time range:   B08012_003E - B08012_013E
    fields_list = ["B08012_0" + ("0" *  (2 -len(str(i)))) + str(i) + "E" for i in range(2,14)]

    
    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            travel_time_raw_data = c.acs5.state_place(fields=fields_list,state_fips=state_fips,place=place_fips)[0]
        except Exception as e:
            print(e, 'Problem getting travel time  data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
            travel_time_raw_data = c.acs5.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]

        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            travel_time_raw_data = c.acs5.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
            
            travel_time_raw_data       = c.acs5.zipcode(fields = fields_list, zcta = '*')
            travel_time_raw_data       = FindZipCodeDictionary(zip_code_data_dictionary_list =   travel_time_raw_data  , zcta = zcta, state_fips = state_fips )

        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
    
    elif geographic_level == 'tract':
        try:
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
            
            travel_time_raw_data = c.acs5.state_county_tract(fields=fields_list,state_fips=state_fips,county_fips=county_fips,tract=tract)[0]
        except Exception as e:
            print(e, 'Problem getting year built data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        travel_time_raw_data = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )





    #Create an empty list and place the values from the dictionary inside of it
    travel_time_data = []
    for field in fields_list:
        travel_time_data.append(travel_time_raw_data[field])

    #Convert list with raw totals into a list where each element is a fraction of the total
    travel_time_data = ConvertListElementsToFractionOfTotal(travel_time_data)
    return(travel_time_data)

#Travel Method to work
def GetTravelMethodData(geographic_level,hood_or_comparison_area):
    print('Getting travel method to work data for: ')
    
    fields_list = ['B08006_001E','B08006_003E','B08006_004E','B08006_015E','B08006_008E','B08006_017E','B08006_016E','B08006_014E']

    if geographic_level == 'place':
        try:
            if hood_or_comparison_area == 'hood':
                place_fips = hood_place_fips

            elif hood_or_comparison_area == 'comparison area':
                place_fips = comparsion_place_fips

            neighborhood_method_to_work_distribution_raw   = c.acs5.state_place(fields=fields_list,state_fips=state_fips,place=place_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting travel method data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county': 
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips

            neighborhood_method_to_work_distribution_raw   = c.acs5.state_county(fields=fields_list,state_fips=state_fips,county_fips=county_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting travel method data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'county subdivision':
        try:
            if hood_or_comparison_area == 'hood':
                county_fips = hood_county_fips
                subdiv_fips = hood_suvdiv_fips

            elif hood_or_comparison_area == 'comparison area':
                county_fips = comparison_county_fips
                subdiv_fips = comparison_suvdiv_fips
            
            neighborhood_method_to_work_distribution_raw   = c.acs5.state_county_subdivision(fields=fields_list,state_fips=state_fips,county_fips=county_fips,subdiv_fips=subdiv_fips)[0]
        
        except Exception as e:
            print(e, 'Problem getting travel method data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()
        
    elif geographic_level == 'zip':
        try:
            if hood_or_comparison_area == 'hood':
                zcta = hood_zip

            elif hood_or_comparison_area == 'comparison area':
                zcta = comparison_zip
            
            
            neighborhood_method_to_work_distribution_raw       = c.acs5.zipcode(fields = fields_list, zcta = '*')
            neighborhood_method_to_work_distribution_raw       = FindZipCodeDictionary(zip_code_data_dictionary_list =   neighborhood_method_to_work_distribution_raw  , zcta = zcta, state_fips = state_fips )
        
        except Exception as e:
            print(e, 'Problem getting travel method data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'tract':
        try:        
            if hood_or_comparison_area == 'hood':
                tract       = hood_tract 
                county_fips = hood_county_fips

            elif hood_or_comparison_area == 'comparison area':
                tract       = comparison_tract
                county_fips = comparison_county_fips
            
            neighborhood_method_to_work_distribution_raw   = c.acs5.state_county_tract(fields=fields_list,state_fips=state_fips,county_fips=county_fips,tract=tract)[0]
        
        except Exception as e:
            print(e, 'Problem getting travel method data for: Geographic Level - ' + geographic_level + ' for ' + hood_or_comparison_area )
            return()

    elif geographic_level == 'custom':
        #Create empty list we will fill with dictionaries (one for each census tract within the custom shape/neighborhood)
        neighborhood_tracts_data = []

        #Fetch census data for all relevant census tracts within the neighborhood
        raw_census_data = c_area.acs5.geo_tract(fields_list, neighborhood_shape)
        
        for tract_geojson, tract_data, tract_proportion in raw_census_data:
            neighborhood_tracts_data.append((tract_data))

        #Convert the list of dictionaries into a single dictionary where we aggregate all values across keys
        neighborhood_method_to_work_distribution_raw = AggregateAcrossDictionaries(neighborhood_tracts_data = neighborhood_tracts_data, fields_list = fields_list )

    neighborhood_method_to_work_distribution = []
    for field in fields_list:
        neighborhood_method_to_work_distribution.append(neighborhood_method_to_work_distribution_raw[field])
        

    neighborhood_method_to_work_distribution = ConvertListElementsToFractionOfTotal(neighborhood_method_to_work_distribution)
        
    return(neighborhood_method_to_work_distribution) 

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

        _2010_hood_pop         = c.sf1.state_place(fields = total_pop_field,                                        state_fips = state_fips, place = hood_place_fips)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_place(fields = total_households_field,                                 state_fips = state_fips, place = hood_place_fips)[0][total_households_field]
        
        current_hood_pop       = c.pl.state_place(fields = [redistricting_total_pop_field],                          state_fips = state_fips, place = hood_place_fips)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_place(fields = [redistricting_total_hh_field],                           state_fips = state_fips, place = hood_place_fips)[0][redistricting_total_hh_field]
        
    elif hood_geographic_level == 'county':
        current_estimate_period = '2020 Census'

        _2010_hood_pop   = c.sf1.state_county(fields = total_pop_field,                      state_fips = state_fips, county_fips = hood_county_fips)[0][total_pop_field]
        _2010_hood_hh    = c.sf1.state_county(fields = total_households_field,               state_fips = state_fips, county_fips = hood_county_fips)[0][total_households_field]
    
        current_hood_pop =  c.pl.state_county(fields = redistricting_total_pop_field,        state_fips = state_fips, county_fips = hood_county_fips)[0][redistricting_total_pop_field]
        current_hood_hh  =  c.pl.state_county(fields = redistricting_total_hh_field,         state_fips = state_fips, county_fips = hood_county_fips)[0][redistricting_total_hh_field]
        
    elif hood_geographic_level == 'county subdivision':
        current_estimate_period = '2020 Census'
        _2010_hood_pop         = c.sf1.state_county_subdivision(fields = total_pop_field,                     state_fips = state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_county_subdivision(fields = total_households_field,              state_fips = state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][total_households_field]

        current_hood_pop       = c.pl.state_county_subdivision(fields = redistricting_total_pop_field,        state_fips = state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_county_subdivision(fields = redistricting_total_hh_field,         state_fips = state_fips, county_fips = hood_county_fips, subdiv_fips = hood_suvdiv_fips)[0][redistricting_total_hh_field]

    elif hood_geographic_level == 'zip':
        current_estimate_period = 'Current Estimate'

        _2010_hood_pop         = c.sf1.state_zipcode(fields = total_pop_field,        state_fips = state_fips, zcta = hood_zip)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_zipcode(fields = total_households_field, state_fips = state_fips, zcta = hood_zip)[0][total_households_field]

        current_hood_pop       = _2010_hood_pop
        current_hood_hh        = _2010_hood_hh

    elif hood_geographic_level == 'tract':
        current_estimate_period = '2020 Census'
        _2010_hood_pop         = c.sf1.state_county_tract(fields = total_pop_field,              state_fips = state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][total_pop_field]
        _2010_hood_hh          = c.sf1.state_county_tract(fields = total_households_field,       state_fips = state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][total_households_field]

        current_hood_pop       = c.pl.state_county_tract(fields = redistricting_total_pop_field, state_fips = state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][redistricting_total_pop_field]
        current_hood_hh        = c.pl.state_county_tract(fields = redistricting_total_hh_field,  state_fips = state_fips,county_fips=hood_county_fips,tract=hood_tract)[0][redistricting_total_hh_field]

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
        _2010_comparison_pop = c.sf1.state_place(fields = total_pop_field,                       state_fips = state_fips, place = comparsion_place_fips)[0][total_pop_field]
        _2010_comparison_hh  = c.sf1.state_place(fields = total_households_field,                state_fips = state_fips, place = comparsion_place_fips)[0][total_households_field]

        current_comparison_pop = c.pl.state_place(fields = redistricting_total_pop_field,        state_fips = state_fips, place = comparsion_place_fips)[0][redistricting_total_pop_field]
        current_comparison_hh  = c.pl.state_place(fields = redistricting_total_hh_field,         state_fips = state_fips, place = comparsion_place_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'county':
        _2010_comparison_pop   = c.sf1.state_county(fields = total_pop_field,                      state_fips = state_fips, county_fips = comparison_county_fips)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_county(fields = total_households_field,               state_fips = state_fips, county_fips = comparison_county_fips)[0][total_households_field]

        current_comparison_pop =  c.pl.state_county(fields = redistricting_total_pop_field,        state_fips = state_fips, county_fips = comparison_county_fips)[0][redistricting_total_pop_field]
        current_comparison_hh  =  c.pl.state_county(fields = redistricting_total_hh_field,         state_fips = state_fips, county_fips = comparison_county_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'county subdivision':
        _2010_comparison_pop    = c.sf1.state_county_subdivision(fields = total_pop_field,                     state_fips = state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][total_pop_field]
        _2010_comparison_hh     = c.sf1.state_county_subdivision(fields = total_households_field,              state_fips = state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][total_households_field]

        current_comparison_pop  = c.pl.state_county_subdivision(fields = redistricting_total_pop_field,        state_fips = state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][redistricting_total_pop_field]
        current_comparison_hh   = c.pl.state_county_subdivision(fields = redistricting_total_hh_field,         state_fips = state_fips, county_fips = comparison_county_fips, subdiv_fips = comparison_suvdiv_fips)[0][redistricting_total_hh_field]

    elif comparison_geographic_level == 'zip':
        _2010_comparison_pop   = c.sf1.state_zipcode(fields = total_pop_field,state_fips=state_fips,zcta = comparison_zip)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_zipcode(fields = total_households_field,state_fips=state_fips,zcta=comparison_zip)[0][total_households_field]

        current_comparison_pop = _2010_comparison_pop
        current_comparison_hh  = _2010_comparison_hh

    elif comparison_geographic_level == 'tract':
        _2010_comparison_pop   = c.sf1.state_county_tract(fields = total_pop_field,                     state_fips = state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][total_pop_field]
        _2010_comparison_hh    = c.sf1.state_county_tract(fields = total_households_field,              state_fips = state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][total_households_field]

        current_comparison_pop =  c.pl.state_county_tract(fields = redistricting_total_pop_field,        state_fips = state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][redistricting_total_pop_field]
        current_comparison_hh  =  c.pl.state_county_tract(fields = redistricting_total_hh_field,         state_fips = state_fips, county_fips = comparison_county_fips, tract = comparison_tract)[0][redistricting_total_hh_field]

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

#Non Census Sources
def GetWikipediaPage():
    global page
    if (neighborhood_level == 'place') or (neighborhood_level == 'county subdivision') or (neighborhood_level == 'county'): #Don't bother looking for wikipedia page if zip code
        try:
            wikipedia_page_search_term    = (neighborhood + ', ' + state)
            page                          =  wikipedia.page(wikipedia_page_search_term)
                
        except Exception as e:
            print(e,': problem getting wikipedia page')
            
    elif (neighborhood_level == 'custom'):
        try:
            wikipedia_page_search_term    = (neighborhood + ', ' + comparison_area +', ' + state)
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
            search_term = (neighborhood + ', ' + state)

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
        search_term = 'https://www.apartments.com/' + '-'.join(neighborhood.lower().split(' ')) + '-' + state.lower() + '/'
    elif neighborhood_level == 'custom':
        search_term = 'https://www.apartments.com/' + '-'.join(neighborhood.lower().split(' ')) + '-' + '-'.join(comparison_area.lower().split(' ')) +  '-' + state.lower() + '/'


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
            descriptive_paragraphs.append(paragraph.text)
        
        return(descriptive_paragraphs)
    
    
    
    except Exception as e:
        print(e)
        return([''])

def LocationIQ(lat,lon,radius):
    #Searches the Locate IQ API for points of interest
    print('Searching Location IQ API')

    url = "https://us1.locationiq.com/v1/nearby.php"

    data = {
    'key': location_iq_api_key,
    'lat': lat,
    'lon': lon,
    'tag': 'all',
    'radius': radius,
    'format': 'json'
        }

    try:
        response = requests.get(url, params=data).json()
    except Exception as e:
        print(e)
        response = [{}]
        
    # for poi in response:
    #     print(poi['name'],poi['type'],)
    #     print('------------')
        

    return(response)

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

    print('Getting Data for ' + neighborhood)
    neighborhood_household_size_distribution     = GetHouseholdSizeData(     geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Neighborhood households by size
    neighborhood_tenure_distribution             = GetHousingTenureData(     geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Tenure (owner occupied/renter)
    neighborhood_housing_value_data              = GetHousingValues(         geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Owner Occupied housing units by value
    neighborhood_number_units_data               = GetNumberUnitsData(       geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Units by units in building
    neighborhood_year_built_data                 = GetHouseYearBuiltData(    geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Housing Units by year structure built
    neighborhood_method_to_work_distribution     = GetTravelMethodData(      geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Travel Mode to Work
    neighborhood_household_income_data           = GetHouseholdIncomeValues( geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Households by household income data
    neighborhood_age_data                        = GetAgeData(               geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Population by age data
    neighborhood_time_to_work_distribution       = GetTravelTimeData(        geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Travel Time to Work
    neighborhood_top_occupations_data            = GetTopOccupationsData(    geographic_level = neighborhood_level, hood_or_comparison_area = 'hood')          #Top Employment Occupations
    
    print('Getting Data For ' + comparison_area)
    comparison_household_size_distribution       = GetHouseholdSizeData(    geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_tenure_distribution               = GetHousingTenureData(    geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_housing_value_data                = GetHousingValues(        geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')    
    comparison_number_units_data                 = GetNumberUnitsData(      geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')    
    comparison_year_built_data                   = GetHouseYearBuiltData(   geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_age_data                          = GetAgeData(              geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_household_income_data             = GetHouseholdIncomeValues(geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')   
    comparison_top_occupations_data              = GetTopOccupationsData(   geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    comparison_time_to_work_distribution         = GetTravelTimeData(       geographic_level  = comparison_level,   hood_or_comparison_area = 'comparison area')
    
    #Walk score
    walk_score_data                              = GetWalkScore(            lat = latitude, lon = longitude                                                    )

    #Overview Table Data
    overview_table_data = GetOverviewTable(hood_geographic_level = neighborhood_level ,comparison_geographic_level = comparison_level )
    
    #Unused functions
    #SearchGreatSchoolDotOrg()
    
#Graph Related Functions
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
    
    occupations_categories = list(neighborhood_top_occupations_data.keys())
    neighborhood_top_occupations = list(neighborhood_top_occupations_data.values())

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

#Langauge Related Functions    
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
                return(neighborhood + ' is not served by any commuter or light rail lines.')
            else:
                return('')

    except Exception as e:
        print(e,'problem getting wikipedia language for ' + category)
        return('')

def SummaryLangauge():
    try:
        wikipedia_summary = (wikipedia.summary((neighborhood + ',' + state)))
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

    community_assets_language = (neighborhood + ' has a number of community assets. ')
    
    location_iq_data                             = LocationIQ(              lat = latitude, lon = longitude, radius = 5000                                     )
    
    for poi in location_iq_data:
        poi_name      = poi['display_name']
        poi_type      = poi['type']
        poi_city      = poi['address']['city']
        poi_sentence  = (' ' + poi_name + ' is a ' + poi_type + '.')
        community_assets_language = community_assets_language + poi_sentence

    # print(poi)
    print(location_iq_data)
    # print(type(location_iq_data))

    return([community_assets_language])
    
    
  


    # try:
    #     yelp_language = YelpLanguage(yelp_data)
    # except Exception as e:
    #     print(e)
    #     yelp_language =''

    #     #Takes a dictionary as input and returns string
    # assert type(yelp_data) == dict

    # return_string = ''
    # for category in yelp_data.keys(): 
    #     yelp_string = category.title() + ': ' + ', '.join(yelp_data[category]) + '. '
    #     return_string = return_string + yelp_string

    # return(return_string)

    # global yelp_data,location_iq_data
    
    # yelp_data                                    = GetYelpData(             lat = latitude, lon = longitude, radius = 30000                                    ) #radius in meters

def CarLanguage():
    wikipedia_car_language     = WikipediaTransitLanguage(category='car')
    
    if wikipedia_car_language == '':
        return(FindNearestHighways(lat = latitude, lon = longitude))
    else:
        return(wikipedia_car_language)

def PlaneLanguage():
    wikipedia_plane_language     = WikipediaTransitLanguage(category='air')
    nearest_airport_language     = FindNearestAirport(lat = latitude, lon = longitude)

    if wikipedia_plane_language == '':
        return(nearest_airport_language)
    else:
        return(wikipedia_plane_language)

def BusLanguage():
    wikipedia_bus_language = WikipediaTransitLanguage(category='bus')
    return(wikipedia_bus_language)

def TrainLanguage():
    wikipedia_train_language = WikipediaTransitLanguage(category='train')
    return(wikipedia_train_language)

def OutlookLanguage():
    return('Neighborhood analysis can best be summarized by referring to neighborhood life cycles. ' +
          'Neighborhoods are perceived to go through four cycles, the first being growth, the second being stability, the third decline, and the fourth revitalization. ' +
          'It is our observation that the subject’s neighborhood is exhibiting several stages of the economic life, with an overall predominance of stability and both limited decline and limited revitalization in some sectors. ' +
          'The immediate area surrounding the subject, has had a historically low vacancy level and is located just to the south of the ------ submarket,' +
          """ which has multiple office and retail projects completed within the past two years and more development in the subject’s immediate vicinity either under construction or preparing to break ground."""+
          ' The proximity of the ________ and ________ will ensure the neighborhood will continue ' +
          'to attract growth in the long-term.')
    pass

def TransportationOverviewLanguage():
    try:
        transportation_language         =  page.section('Transportation')
    except Exception as e:
        print(e)
        transportation_language         = ''
    
    return(transportation_language)

def HousingTenureLanguage():
    hood_owner_occupied_fraction        =  neighborhood_tenure_distribution[1] 
    comparsion_owner_occupied_fraction  =  comparison_tenure_distribution[1]

    if hood_owner_occupied_fraction > comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'higher than'
    
    elif hood_owner_occupied_fraction < comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'lower than'

    elif hood_owner_occupied_fraction > comparsion_owner_occupied_fraction:
        hood_owner_ouccupied_higher_lower   =  'equal to'
            
    else:
        hood_owner_ouccupied_higher_lower   =  '[lower than/higher than/equal to]'
    
    tenure_langugage = ('According to the most recent American Community Survey, ' +
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
        #    '. This chart shows the ownership percentage in ' + neighborhood + ' compared to _______. ')

    return([tenure_langugage])

def HousingSizeLanguage():
    number_units_categories = ['Single Family Homes','Townhomes','Duplexes','3-4 Units','5-9 Units','10-19 Units','20-49 Units','50 >= Units']
    
    most_common_category           = number_units_categories[neighborhood_number_units_data.index(max(neighborhood_number_units_data))] #get the most common number units category


    housing_size_langauge = ( neighborhood + 
                              ' offers a wide variety of housing options including single family homes, some duplexes and smaller multifamily properties, along with larger garden style properties, and even some buildings with 50+ units. ' + 
                              most_common_category +
                              ' is the most common form of housing in ' +
                              neighborhood 
                            
                             )
    return[housing_size_langauge]

def HousingValueLanguage():
  
    housing_value_categories = ['$10,000 <','$10,000-14,999','$15,000-19,999','$20,000-24,999','$25,000-29,999','$30,000-34,000','$35,000-39,999','$40,000-49,000','$50,000-59,9999','$60,000-69,999','$70,000-79,999','$80,000-89,999','$90,000-99,999','$100,000-124,999','$125,000-149,999','$150,000-174,999','$175,000-199,999','$200,000-249,999','$250,000-299,999','$300,000-399,999','$400,000-499,999','$500,000-749,999','$750,000-999,999','$1,000,000-1,499,999','$1,500,000-1,999,999','$2,000,000 >=']


    #Estimate a median household income from a category freqeuncy distribution
    total_value_fraction = 0
    for i,value_category_fraction in enumerate(neighborhood_housing_value_data):
        total_value_fraction += value_category_fraction
        if total_value_fraction >= 50:
            median_cat_index = i
            break
    
    hood_median_value_range     = housing_value_categories[median_cat_index]
    hood_median_value_range     = hood_median_value_range.replace('$','')
    hood_median_value_range     = hood_median_value_range.replace(',','').split('-')
    hood_median_value           = round((int(hood_median_value_range[0]) + int(hood_median_value_range[1]))/2,1)
    
    hood_largest_value_category = housing_value_categories[neighborhood_housing_value_data.index(max(neighborhood_housing_value_data))] #get the most common income category
    comp_largest_value_category = housing_value_categories[comparison_housing_value_data.index(max(comparison_housing_value_data))]

    value_language = (  'Homes in '                                        +
                       neighborhood                                           + 
                       ' have a median value of about '                      + 
                        "${:,.0f}".format(hood_median_value)                   +
                       ', displayed in the chart below. '                     +
                       
                       'In '                                                  + 
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
    
    year_built_categories = ['2014 >=','2010-2013','2000-2009','1990-1999','1980-1989','1970-1979','1960-1969','1950-1959','1940-1949','<= 1939']
    year_built_categories.reverse()

    #Estimate a median household income from a category freqeuncy distribution
    total_yrblt_fraction = 0
    for i,yrblt_category_fraction in enumerate(neighborhood_year_built_data):
        total_yrblt_fraction += yrblt_category_fraction
        if total_yrblt_fraction >= 50:
            median_cat_index = i
            break
    
    hood_median_yrblt_range     = year_built_categories[median_cat_index]
    hood_median_yrblt_range     = hood_median_yrblt_range.replace('<= ', '')
    if len(hood_median_yrblt_range) == 4:
        hood_median_yrblt = int(hood_median_yrblt_range)
    else:
        hood_median_yrblt_range     = hood_median_yrblt_range.replace(' >=', '').split('-')
        hood_median_yrblt           = round((int(hood_median_yrblt_range[0]) + int(hood_median_yrblt_range[1]))/2,1)
    
    hood_largest_yrblt_category = year_built_categories[neighborhood_year_built_data.index(max(neighborhood_year_built_data))] #get the most common income category
    comp_largest_yrblt_category = year_built_categories[comparison_year_built_data.index(max(comparison_year_built_data))]

    yrblt_language = (  'Homes in '                                         +
                       neighborhood                                         + 
                       ' have a median year built of about '                + 
                        "{:.0f}".format(hood_median_yrblt)                 +
                       ', displayed in the chart below. '                   +
                       
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
    return(['The majority of working age residents are employed in the ______, ______, and _______ industries. ' + 'Given its large share of employees in the production industry, '])

def HouseholdSizeLanguage():
    
    household_size_categories = ['1','2','3','4','5','6','7+']


    #Estimate a median household size from a category freqeuncy distribution
    total_size_fraction = 0
    for i,size_category_fraction in enumerate(neighborhood_household_size_distribution):
        total_size_fraction += size_category_fraction
        if total_size_fraction >= 50:
            median_cat_index = i
            break
    
    hood_median_size   = int(household_size_categories[median_cat_index].replace('+',''))
  

    hood_largest_time_category = household_size_categories[neighborhood_household_size_distribution.index(max(neighborhood_household_size_distribution))] #get the most common household size category
    comp_largest_time_category = household_size_categories[comparison_household_size_distribution.index(max(comparison_household_size_distribution))]

    household_size_language = ('Households in '                                        +
                               neighborhood                                           + 
                              ' have a median size of '                      + 
                        "{:,.0f} people".format(hood_median_size)                   +
                       ', displayed in the chart below. '                     +
                       
                       'In '                                                  + 
                       neighborhood                                           + 
                       ', the largest share of households have ' +
                       hood_largest_time_category +
                       ' people, compared to ' +
                       comp_largest_time_category        +
                       ' for '           +
                        comparison_area +
                        '.'
                        )
    
    return([household_size_language])

def PopulationAgeLanguage():
    age_ranges = ['0-19','20-24','25-34','35-49','50-66','67+']

    #Estimate a median age  from a category freqeuncy distribution
    total_pop_fraction = 0
    for i,age_category_fraction in enumerate(neighborhood_age_data):
        total_pop_fraction += age_category_fraction
        if total_pop_fraction >= 50:
            median_cat_index = i
            break

    hood_median_age_range      = age_ranges[median_cat_index]
    hood_median_age_range      = hood_median_age_range.replace(',','').split('-')
    hood_median_age            = round((int(hood_median_age_range[0]) + int(hood_median_age_range[1]))/2,1)
    
    hood_largest_age_category = age_ranges[neighborhood_age_data.index(max(neighborhood_age_data))] #get the most common income category
    comp_largest_age_category = age_ranges[comparison_age_data.index(max(comparison_age_data))]

    age_language = ('The median age in '                                                        +
                       neighborhood                                                         + 
                       ' is around '                         + 
                        "{:,.0f}".format(hood_median_age)                                   +
                       ', displayed in the chart below. '                                   +
                       
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

    # age_language = ('The population in ' + neighborhood + ' has increased/compressed X% to ___ as of date. ' + 'Given a median age of _____, and a large distribution towards middle aged individuals, family households with school-aged children are most common.')
    
    return([age_language])

def IncomeLanguage():

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
    total_inc_fraction = 0
    for i,income_category_fraction in enumerate(neighborhood_household_income_data):
        total_inc_fraction += income_category_fraction
        if total_inc_fraction >= 50:
            median_cat_index = i
            break

    hood_median_income_range   = income_categories[median_cat_index]
    hood_median_income_range   = hood_median_income_range.replace('$','')
    hood_median_income_range   = hood_median_income_range.replace(',','').split('-')
    
    hood_median_income         = round((int(hood_median_income_range[0]) + int(hood_median_income_range[1]))/2,1)
    
    hood_largest_income_category = income_categories[neighborhood_household_income_data.index(max(neighborhood_household_income_data))] #get the most common income category
    comp_largest_income_category = income_categories[comparison_household_income_data.index(max(comparison_household_income_data))]

    income_language = ('Households in '                                      +
                       neighborhood                                          + 
                       ' have a median household income of around '          + 
                        "${:,.0f}".format(hood_median_income)                 +
                       ', displayed in the chart below. '                    +
                       
                       'In '                                                 + 
                       neighborhood                                          + 
                       ', the largest share of households have a household income between ' +
                       hood_largest_income_category +
                       ', compared to ' +
                       comp_largest_income_category        +
                       ' for '           +
                        comparison_area +
                        '.'
                        )
    
    return([income_language])

def TravelMethodLanguage():
    travel_method_categories = ['driving alone','car pooling','public transportation','walking','working from home','biking','other']
    hood_largest_travel_category      = travel_method_categories[neighborhood_method_to_work_distribution.index(max(neighborhood_method_to_work_distribution))] #get the most common income category
    hood_largest_travel_category_frac = neighborhood_method_to_work_distribution[neighborhood_method_to_work_distribution.index(max(neighborhood_method_to_work_distribution))]
    travel_method_language = ('In ' + neighborhood + ', the most common method for traveling to work is ' + hood_largest_travel_category.lower()  + ' with ' +  "{:,.1f}%".format(hood_largest_travel_category_frac) + ' of commuters using it.')
    return([travel_method_language])
    
def TravelTimeLanguage():
    
    travel_time_categories = ['< 5 Minutes','5-9 Minutes','10-14 Minutes','15-19 Minutes','20-24 Minutes','25-29 Minutes','30-34 Minutes','35-39 Minutes','40-44 Minutes','45-59 Minutes','60-89 Minutes','> 90 Minutes']


    #Estimate a median household income from a category freqeuncy distribution
    total_time_fraction = 0
    for i,time_category_fraction in enumerate(neighborhood_time_to_work_distribution):
        total_time_fraction += time_category_fraction
        if total_time_fraction >= 50:
            median_cat_index = i
            break
    
    hood_median_time_range   = travel_time_categories[median_cat_index]
    hood_median_time_range   = hood_median_time_range.replace(' Minutes','')
    hood_median_time_range   = hood_median_time_range.replace(',','').split('-')
    hood_median_time         = round((int(hood_median_time_range[0]) + int(hood_median_time_range[1]))/2,1)
    
    hood_largest_time_category = travel_time_categories[neighborhood_time_to_work_distribution.index(max(neighborhood_time_to_work_distribution))] #get the most common income category
    comp_largest_time_category = travel_time_categories[comparison_time_to_work_distribution.index(max(comparison_time_to_work_distribution))]

    time_language = ('Commuters in '                                        +
                       neighborhood                                           + 
                       ' have a median commute time of about '                      + 
                        "{:,.0f} minutes".format(hood_median_time)                   +
                       ', displayed in the chart below. '                     +
                       
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

def CreateLanguage():
    print('Creating Langauge')

    global bus_language,car_language,plane_language,train_language,transportation_language,summary_langauge,conclusion_langauge,community_assets_language
    global yelp_language
    global airport_language
    global apartmentsdotcomlanguage
    global housing_tenure_breakdown_language, structure_size_breakdown_language
    global employment_language
    global population_age_language
    global income_language
    global travel_method_language, travel_time_language
    global housing_value_language,year_built_language
    global household_size_language

    summary_langauge                   =  SummaryLangauge()
    community_assets_language          =  CommunityAssetsLanguage()
    transportation_language            =  TransportationOverviewLanguage()

    housing_tenure_breakdown_language  = HousingTenureLanguage()
    structure_size_breakdown_language  = HousingSizeLanguage()
    
    housing_value_language             = HousingValueLanguage()
    year_built_language                = HousingYearBuiltLanguage()
    
    employment_language                = EmploymentLanguage()
    population_age_language            = PopulationAgeLanguage()
    income_language                    = IncomeLanguage()
    household_size_language            = HouseholdSizeLanguage()

    travel_method_language             = TravelMethodLanguage() 
    travel_time_language               = TravelTimeLanguage()


    bus_language                       = BusLanguage() 
    train_language                     = TrainLanguage()
    car_language                       = CarLanguage()
    plane_language                     = PlaneLanguage()

    conclusion_langauge                = OutlookLanguage()
    
#Report document related functions
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

    glance_paragraph                               = document.add_paragraph(neighborhood + ' at a Glance')
    glance_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
    glance_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

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

def Citation(document,text):
    citation_paragraph = document.add_paragraph()
    citation_paragraph.paragraph_format.space_after  = Pt(6)
    citation_paragraph.paragraph_format.space_before = Pt(6)
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
    try:
        #Search Google Maps for hood
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        browser = webdriver.Chrome(executable_path=(os.path.join(os.environ['USERPROFILE'], 'Desktop','chromedriver.exe')),options=options)
        browser.get('https:google.com/maps')
            
        #Write hood name in box
        Place = browser.find_element_by_class_name("tactile-searchbox-input")

        if neighborhood_level != 'custom':
            Place.send_keys((neighborhood + ', ' + state))
        
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
        time.sleep(1)

    
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
    
    #Add image of map
    map_path  =  os.path.join(hood_folder_map,'map.png')
    map2_path = os.path.join(hood_folder_map,'map2.png')
    map3_path = os.path.join(hood_folder_map,'map3.png')

    #Open zommed out map
    img1 = Image.open(map2_path)
    
    #Open zommed in map

    add_border(map_path, output_image = map_path, border=1)
    time.sleep(.2)
    img2 = Image.open(map_path)

    #Reduce size of zommed in image by a constant factor
    image_reduction_scale = 3
    img2 = img2.resize((int(img2.size[0]/image_reduction_scale),int(img2.size[1]/image_reduction_scale)))
    
    #Add the zoomed in map on top of the zoomed out map and save as new png image
    # No transparency mask specified,                                      
    # simulating an raster overlay
    img1.paste(img2, (img1.size[1] - 25,900))
    
    # img1.show()
    img1.save(map3_path)
    
def AddMap(document):
    print('Adding Map')
    map_path  = os.path.join(hood_folder_map,'map.png')
    map2_path = os.path.join(hood_folder_map,'map2.png')
    map3_path = os.path.join(hood_folder_map,'map3.png')
    
    if os.path.exists(map_path) == False or os.path.exists(map2_path) == False: #If we don't have a zommed in map image or a zoomed out map, create one
        GetMap()    

    if os.path.exists(map3_path) == False: #If we don't have an image with a zommed in map overlayed on zoomed out map, create one
        OverlayMapImages()
    
    if os.path.exists(map3_path):
        map = document.add_picture(map3_path,width=Inches(6.5))

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

                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(8)
                    run.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    #make first row bold
                    if current_row == 0: 
                        font.bold = True
                        font.name = 'Avenir Next LT Pro Demi'

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
    AddTitle(document = document)
    AddMap(document = document)
    Citation(document,'Google Maps')
    AddHeading(document = document, title = 'Overview',            heading_level = 1,heading_number='Heading 3',font_size=11)
    
    #Add neighborhood overview language
    for paragraph in summary_langauge:
        if paragraph == '':
            continue
        summary_paragraph                               = document.add_paragraph(paragraph)
        summary_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        summary_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)
        summary_format                                  = document.styles['Normal'].paragraph_format
        summary_format.line_spacing_rule                = WD_LINE_SPACING.SINGLE
        summary_style                                   = summary_paragraph.style
        summary_style.font.name                         = primary_font

    try:
        #Add Overview Table
        AddTable(document = document,data_for_table = overview_table_data )
    except Exception as e:
        print(e,'Unable to add overview table')

def CommunityAssetsSection(document):
    print('Writing Community Assets Section')
    #Community Assets Section
    AddHeading(document = document, title = 'Community Assets',            heading_level = 1,heading_number='Heading 3',font_size=11)

    for paragraph in community_assets_language:
        if paragraph == '':
            continue
        community_paragraph                               = document.add_paragraph(paragraph)
        community_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        community_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

def HousingSection(document):
    print('Writing Neighborhood Section')
    AddHeading(document = document, title = 'Housing',                  heading_level = 1,heading_number='Heading 3',font_size=11)
    
    #Add structure size language
    for paragraph in structure_size_breakdown_language:
        if paragraph == '':
            continue
        structure_size_paragraph                               = document.add_paragraph(paragraph)
        structure_size_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        structure_size_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

    #Insert household units by units in_structure graph
    if os.path.exists(os.path.join(hood_folder,'household_units_in_structure_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_units_in_structure_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Add tenure language
    for paragraph in housing_tenure_breakdown_language:
        if paragraph == '':
            continue
        tenure_paragraph                               = document.add_paragraph(paragraph)
        tenure_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        tenure_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

    #Insert Household Tenure graph
    if os.path.exists(os.path.join(hood_folder,'household_tenure_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_tenure_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Add housing value language
    for paragraph in housing_value_language:
        if paragraph == '':
            continue
        housing_value_paragraph                               = document.add_paragraph(paragraph)
        housing_value_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        housing_value_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

    #Insert Household value graph
    if os.path.exists(os.path.join(hood_folder,'household_value_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_value_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Add language
    for paragraph in year_built_language :
        if paragraph == '':
            continue
        year_built_paragraph                                   = document.add_paragraph(paragraph)
        year_built_paragraph.alignment                         = WD_ALIGN_PARAGRAPH.JUSTIFY
        year_built_paragraph.paragraph_format.space_after      = Pt(primary_space_after_paragraph)



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
    for paragraph in household_size_language:
        if paragraph == '':
            continue
        hh_size_paragraph                               = document.add_paragraph(paragraph)
        hh_size_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        hh_size_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)


    #Insert Household size graph
    if os.path.exists(os.path.join(hood_folder,'household_size_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'household_size_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')

    #Age langauge
    for paragraph in population_age_language:
        if paragraph == '':
            continue
        pop_paragraph                               = document.add_paragraph(paragraph)
        pop_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        pop_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

    #Insert population by age graph
    if os.path.exists(os.path.join(hood_folder,'population_by_age_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'population_by_age_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    



    #Income langauge
    for paragraph in income_language:
        if paragraph == '':
            continue
        inc_paragraph                               = document.add_paragraph(paragraph)
        inc_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        inc_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

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

    for paragraph in employment_language:
        if paragraph == '':
            continue
        emp_paragraph                               = document.add_paragraph(paragraph)
        emp_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        emp_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

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
    for paragraph in travel_time_language:
        if paragraph == '':
            continue
        travel_time_paragraph                               = document.add_paragraph(paragraph)
        travel_time_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        travel_time_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)

    #Insert Travel Time to Work graph
    if os.path.exists(os.path.join(hood_folder,'travel_time_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'travel_time_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Travel method Lanaguage
    for paragraph in travel_method_language:
        if paragraph == '':
            continue
        travel_time_paragraph                               = document.add_paragraph(paragraph)
        travel_time_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
        travel_time_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)
    
    #Insert Transport Method to Work graph
    if os.path.exists(os.path.join(hood_folder,'travel_mode_graph.png')):
        fig = document.add_picture(os.path.join(hood_folder,'travel_mode_graph.png'),width=Inches(fig_width))
        last_paragraph = document.paragraphs[-1] 
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        Citation(document,'U.S. Census Bureau')
    
    #Transportation Methods table
    table_paragraph             = document.add_paragraph('Transportation Methods')
    table_paragraph.alignment   = WD_ALIGN_PARAGRAPH.CENTER
    transportation_paragraph    = document.add_paragraph(transportation_language)

    #Insert the transit graphics(car, bus,plane, train)
    tab = document.add_table(rows=1, cols=2)
    for pic in ['car.png','train.png','bus.png','plane.png']:
        row_cells = tab.add_row().cells
        paragraph = row_cells[0].paragraphs[0]
        run = paragraph.add_run()
        if pic == 'car.png':
            run.add_text(' ')
        run.add_picture(os.path.join(graphics_location,pic))
    


    transit_table_language = [car_language, train_language, bus_language, plane_language]

    #Loop through the rows in the table
    for current_row ,row in enumerate(tab.rows): 
        #loop through all cells in the current row
        for current_column,cell in enumerate(row.cells):
            if current_column == 1 and current_row > 0:
                cell.text = str(transit_table_language[current_row-1])

            if current_column == 0:
                cell.width = Inches(.2)
            else:
                cell.width = Inches(6)







    #Walk/Bike/Transit Score Table
    table_paragraph = document.add_paragraph('Walk, Bike, and Transit Scores')
    table_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #Add transit score table
    tab = document.add_table(rows=1, cols=2)
    for pic in ['car.png','car.png','car.png',]:
        row_cells = tab.add_row().cells
        paragraph = row_cells[0].paragraphs[0]
        run = paragraph.add_run()
        if pic == 'car.png':
            run.add_text(' ')
        run.add_picture(os.path.join(graphics_location,pic))
    


  
    #Loop through the rows in the table
    for current_row ,row in enumerate(tab.rows): 
        #loop through all cells in the current row
        for current_column,cell in enumerate(row.cells):
            if current_column == 1 and current_row > 0:
                try:
                    cell.text = str(walk_score_data[current_row-1])
                except:
                    cell.text = 'NA'


            if current_column == 0:
                cell.width = Inches(.2)
            else:
                cell.width = Inches(6)
    Citation(document,'https://www.walkscore.com/')

def OutlookSection(document):
    print('Writing Outlook Section')
    AddHeading(document = document, title = 'Conclusion',            heading_level = 1,heading_number='Heading 3',font_size=11)
    conclusion_paragraph                               = document.add_paragraph(conclusion_langauge)
    conclusion_paragraph.alignment                     = WD_ALIGN_PARAGRAPH.JUSTIFY
    conclusion_paragraph.paragraph_format.space_after  = Pt(primary_space_after_paragraph)
    conclusion_style                                   = conclusion_paragraph.style
    conclusion_style.font.name                         = primary_font
    
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
    if testing_mode == False:
        #Give the user 10 seconds to decide if writing reports for metro areas or individual county entries
        try:
            # report_creation = input_with_timeout('Create new report? y/n', 10).strip()
            report_creation = 'y'

        except TimeoutExpired:
            report_creation = ''

        # report_creation = 'y'

    else:
        report_creation = 'y'

def UserSelectsNeighborhoodLevel():
    
    # # Get Input from User
    # allowable_area_levels       = ['p','c','sd','t','custom','z']

    # #Ask user for info on subject area
    # while True:
    #     if testing_mode == False:
    #         neighborhood_level = input('What is the geographic level of the neighborhood? (p = place,sd = subdivision, c = county,t = tract,custom = custom, z = zip code)')
    #     else:
    #         neighborhood_level   =  'p'
        
    #     if neighborhood_level not in allowable_area_levels:
    #         print('Not a supported geographic level for neighborhood area')
    #         continue
    #     else:
    #         break
    
    # #Ask user for info on comparison area
    # while True:
    #     if testing_mode == False:
    #         comparison_level   = input('What is the geographic level of the comparison area? (p = place,sd = subdivision, c = county,t = tract,custom = custom, z = zip code)')
    #     else:
    #         comparison_level     = 'c'
        
    #     if comparison_level not in allowable_area_levels:
    #         print('Not a supported geographic level for comparsion area')
    #         continue
    #     else:
    #         break
    
    # return([neighborhood_level,comparison_level])

    # return(['custom','p'])
    return(['p','c'])

def GetUserInputs():
    
    global neighborhood_level,comparison_level
    hood_comparison_levels = UserSelectsNeighborhoodLevel()
    neighborhood_level     = hood_comparison_levels[0] 
    comparison_level       = hood_comparison_levels[1] 
    
    global neighborhood, hood_tract, hood_zip, hood_place_fips, place_type, hood_suvdiv_fips, hood_county_fips
    global state, state_fips, state_full_name
    global comparison_area, comparison_tract ,comparison_zip, comparsion_place_fips, comparison_suvdiv_fips, comparison_county_fips    

    #Get User input on neighborhood/subject area
    if neighborhood_level == 'p':        #when our neighborhood is a town or city eg: East Rockaway Village, New York
        

        neighborhood_level = 'place'
        if testing_mode == False:
            # fips = input('Enter the 7 digit Census Place FIPS Code')
            fips = '44-49960'
        else:
            fips = '36-22876'

        #Process the FIPS code provided    
        fips            = fips.replace('-','').strip()
        state_fips      = fips[0:2]
        hood_place_fips = fips[2:]
        assert len(fips) == 7

        #Get name of the hood using the FIPS code provided
        neighborhood = c.sf1.state_place(fields=['NAME'],state_fips=state_fips,place=hood_place_fips)[0]['NAME']
        state_full_name = neighborhood.split(',')[1].strip()
        neighborhood = neighborhood.split(',')[0].strip()
        place_type   = neighborhood.split(' ')[len(neighborhood.split(' '))-1] #eg: village, city, etc
        neighborhood = ' '.join(neighborhood.split(' ')[0:len(neighborhood.split(' '))-1]).title()
        

        #Name of State
        state = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state = state.abbr
        assert len(state) == 2

    elif neighborhood_level == 'sd':     #when our neighborhood is county subdivison eg: Town of Hempstead, New York (A large town in Nassau County with several villages within it)
        
        #Assign the nei
        neighborhood_level = 'county subdivision'
        
        #Get FIPS from User
        fips = input('Enter the 10 digit county subdivision FIPS Code for the hood')
        
        #Proccess FIPS code provided
        fips = fips.replace('-','').strip()
        assert len(fips) == 10
        state_fips       = fips[0:2]
        hood_county_fips = fips[2:5]
        hood_suvdiv_fips = fips[5:]

        #Get name of hood using the FIPS code provided
        neighborhood = c.sf1.state_county_subdivision(fields=['NAME'],state_fips=state_fips,county_fips=hood_county_fips,subdiv_fips=hood_suvdiv_fips)[0]['NAME']
        state_full_name = neighborhood.split(',')[2].strip()
        neighborhood = neighborhood.split(',')[0].strip().title()
        place_type   = neighborhood.split(' ')[len(neighborhood.split(' '))-1] #eg: village, city, etc
        neighborhood = ' '.join(neighborhood.split(' ')[0:len(neighborhood.split(' '))-1]).title()

        #Name of State
        state = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state = state.abbr
        assert len(state) == 2

    elif neighborhood_level == 't':      #when our neighborhood is a census tract eg: Tract 106.01 in Manhattan

        neighborhood_level = 'tract' 
        fips = input('Enter the 5 digit County FIPS Code for the county the hood tract is in')
        fips = fips.replace('-','').strip()
        assert len(fips) == 5
        state_fips = fips[0:2]
        hood_county_fips = fips[2:]
        
        hood_tract = input('Enter the 6 digit tract FIPS Code for hood')
        assert len(hood_tract) == 6

        #Get name of hood
        neighborhood = c.sf1.state_county_tract(fields=['NAME'],state_fips=state_fips,county_fips=hood_county_fips,tract=hood_tract)[0]['NAME']
        state_full_name = neighborhood.split(',')[2].strip()
        neighborhood = neighborhood.split(',')[0] + ',' +  neighborhood.split(',')[1]
        neighborhood = neighborhood.strip().title()


        #Name of State
        state = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state = state.abbr
        assert len(state) == 2

    elif neighborhood_level == 'z':      #When our neighborhood is a zip code eg: 11563
        
        #Get input from user
        neighborhood_level              = 'zip'
        hood_zip                        = input('Enter the 5 digit zip code for hood')
        
        #Process the zip code provided
        hood_zip                            = str(hood_zip).replace('-','').strip()
        assert len(hood_zip) == 5
        
        #Get the state FIPS code (eg New York: 36)
        zip_county_crosswalk_df            = pd.read_excel(os.path.join(data_location,'Census Area Codes','ZIP_COUNTY_092021.xlsx')) #read in crosswalk file
        zip_county_crosswalk_df['ZIP']     = zip_county_crosswalk_df['ZIP'].astype(str)
        zip_county_crosswalk_df['ZIP']     = zip_county_crosswalk_df['ZIP'].str.zfill(5)
        zip_county_crosswalk_df['COUNTY']  = zip_county_crosswalk_df['COUNTY'].astype(str)
        zip_county_crosswalk_df['COUNTY']  = zip_county_crosswalk_df['COUNTY'].str.zfill(5)

        zip_county_crosswalk_df            = zip_county_crosswalk_df.loc[zip_county_crosswalk_df['ZIP'] == hood_zip]                 #restrict to rows for zip code
        state_fips                         = str(zip_county_crosswalk_df['COUNTY'].iloc[-1])[0:2] #Get state fips from the county fips code (the county the zip code is in)
        assert len(state_fips) == 2
  
        #Get name of hood
        neighborhood                      = c.sf1.state_zipcode(fields=['NAME'],state_fips=state_fips, zcta=hood_zip)[0]['NAME']
        state_full_name                   = neighborhood.split(',')[1].strip()
        neighborhood                      = neighborhood.split(',')[0].replace('ZCTA5','').strip().title() + ' (Zip Code)'
    
        #Name of State
        state                             = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state                             = state.abbr
        assert                            len(state) == 2

    elif neighborhood_level == 'c':      #When our neighborhood is a county eg Nassau County, New York

        neighborhood_level = 'county'
        
        #Get FIPS code input from user
        fips = input('Enter the 5 digit county FIPS Code for the hood')
        
        #Process the FIPS code provided by user
        fips = fips.replace('-','').strip()
        assert len(fips) == 5
        state_fips       = fips[0:2]
        hood_county_fips = fips[2:]

        #Get name of hood
        neighborhood     = c.sf1.state_county(fields=['NAME'],state_fips=state_fips,county_fips=hood_county_fips)[0]['NAME']
        state_full_name  = neighborhood.split(',')[1].strip()
        neighborhood     = neighborhood.split(',')[0].strip().title()

        #Name of State
        state            = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state            = state.abbr

        assert len(state) == 2

    elif neighborhood_level == 'custom': #When our neighborhood is a neighboorhood within a city (eg: Financial District, New York City)

        #Get name of hood
        neighborhood      = input('Enter the name of the custom neighborhood')
        # neighborhood        = 'Gaslamp Quarter'
            
    #Get user input on comparison area
    if comparison_level == 'c':          #When our comparison area is a county eg Nassau County, New York
        
        comparison_level = 'county'

        #Get comparison county FIPS code from user
        if testing_mode == False:
            # comparison_county_fips = input('Enter the 5 digit FIPS code for the comparison county')
            comparison_county_fips = '44005'

        else:
            comparison_county_fips = '36061'
        
        #Process FIPS code provided by user
        comparison_county_fips = comparison_county_fips.replace('-','').strip()
        assert len(comparison_county_fips) == 5

        #Get name of comparison county using the FIPS code provdided
        comparison_area = c.sf1.state_county(fields=['NAME'],state_fips=comparison_county_fips[0:2],county_fips=comparison_county_fips[2:])[0]['NAME']
        comparison_area = comparison_area.split(',')[0].strip().title()
        comparison_county_fips = comparison_county_fips[2:]

    elif comparison_level == 'p':        #when our comparison area is a town or city eg: East Rockaway Village, New York
        
        #Get place FIPS code from user
        comparison_level      = 'place'
        fips                  = input('Enter the 7 digit Census Place FIPS Code for the comparison area')
        # fips                  = '06-66000'

        #Process FIPS code provided by user``
        fips                  = fips.replace('-','').strip()
        comparsion_place_fips = fips[2:]
        state_fips            = fips[0:2]

        assert len(fips) == 7 and len(state_fips) == 2
        
        #Get name of comparison area
        comparison_area       = c.sf1.state_place(fields=['NAME'],state_fips=state_fips,place=comparsion_place_fips)[0]['NAME']
        state_full_name       = comparison_area.split(',')[1].strip()
        comparison_area       = comparison_area.split(',')[0].strip().title()
        comparison_area       = ' '.join(comparison_area.split(' ')[0:len(comparison_area.split(' '))-1]).title()

        state                 = us.states.lookup(state_full_name) #convert the full state name to the 2 letter abbreviation
        state                 = state.abbr
        assert                len(state) == 2

    elif comparison_level == 'sd':       #when our comparison area is county subdivison eg: Town of Hempstead, New York (A large town in Nassau County with several villages within it)
        comparison_level        = 'county subdivision'

        #Get commaprison area FIPS code from user
        fips                    = input('Enter the 10 digit county subdivision FIPS Code for the comparison area')

        #Process the input from the user
        fips                    = fips.replace('-','').strip()
        assert len(fips) == 10
        comparison_county_fips  = fips[2:5]
        comparison_suvdiv_fips  = fips[5:]

        #Get name of hood
        comparison_area         = c.sf1.state_county_subdivision(fields=['NAME'],state_fips=state_fips, county_fips=comparison_county_fips, subdiv_fips=comparison_suvdiv_fips)[0]['NAME']
        comparison_area         = comparison_area.split(',')[0].strip() 
        comparison_area         = comparison_area.split(' ')[0].title()

    elif comparison_level == 'z':        #When our comparison area is a zip code eg: 11563
        
        comparison_level = 'zip'

        #Get comparison area zip code from user
        comparison_zip = input('Enter the 5 digit zip code for the comparison area')

        #Process user provided zip code
        comparison_zip = comparison_zip.replace('-','').strip()
        assert len(comparison_zip) == 5

        #Get name of comapison area 
        comparison_area = c.sf1.state_zipcode(fields=['NAME'],state_fips=state_fips, zcta=comparison_zip)[0]['NAME']
        comparison_area = comparison_area.split(',')[0].replace('ZCTA5','').strip().title() + ' (Zip Code)'
    
    elif comparison_level == 't':        #when our comparison area is a census tract eg: Tract 106.01 in Manhattan
        
        comparison_level = 'tract'
        
        #Get the FIPS code for the county the comparison tract is in
        fips = input('Enter the 5 digit County FIPS Code for comparison area')
        
        #Process the user provided county FIPS code
        fips = fips.replace('-','').strip()
        assert len(fips) == 5
        comparison_county_fips = fips[2:]
        
        #Get the comaprison tract code from the user
        comparison_tract       = input('Enter the 6 digit tract FIPS Code for comparison area')
        assert len(comparison_tract) == 6

        #Get name of comparison area
        comparison_area = c.sf1.state_county_tract(fields=['NAME'],state_fips=state_fips, county_fips=comparison_county_fips,tract=comparison_tract)[0]['NAME']
        comparison_area = comparison_area.split(',')[0] + ',' +  comparison_area.split(',')[1]
        comparison_area = comparison_area.strip().title()

    elif comparison_level == 'custom':   #When our comparison area is a neighboorhood within a city (eg: Financial District, New York City)
        
        #Get name of comparison area
        comparison_area = input('Enter the name of the custom comparison area')

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

def Main():
    DeclareAPIKeys()
    DecideIfWritingReport()
   

    if report_creation == 'y':
        GetUserInputs() #user selects if they want to run report and gives input for report subject
        print('Preparing report for: ' + neighborhood)
        global latitude
        global longitude
        global current_year
        global neighborhood_shape
        coordinates = GetLatandLon()
        latitude    = coordinates[0] 
        longitude   = coordinates[1]

        if neighborhood_level == 'custom':
            neighborhood_shape = GetNeighborhoodShape()

        todays_date  = date.today()
        current_year = str(todays_date.year)
        SetGraphFormatVariables()
        CreateDirectory()
        GetWikipediaPage()
        GetData()
        # CreateGraphs()
        CreateLanguage()
        WriteReport()
        # CleanUpPNGs()
    
    #Crawl through directory and create CSV with all current neighborhood report documents
    CreateDirectoryCSV()

    #Post an update request to the Market Research Docs Service to update the database
    if main_output_location == os.path.join(dropbox_root,'Research','Market Analysis','Neighborhood'): 
        UpdateServiceDb(report_type='neighborhoods', 
                    csv_name=service_api_csv_name, 
                    csv_path=os.path.join(main_output_location, service_api_csv_name),
                    dropbox_dir='https://www.dropbox.com/home/Research/Market Analysis/Neighborhood/')

#This is our main function that calls all other functions we will use
Main()


